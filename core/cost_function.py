"""
cost_function.py
----------------
Two-weight convolution-based cost function for exam schedule quality.

Classes
-------
ColourWeight       — maps difficulty string to integer multiplier
StudentStressCost  — windowed stress cost across all grades
TeacherMarkingCost — windowed teacher marking load cost
TotalCost          — weighted sum of the two components

Schedule format used throughout:
    {(day: int, session: str): list[ExamPaper]}
    day      — 0-indexed exam day position
    session  — "AM" or "PM"
    value    — list of ExamPaper objects scheduled at that slot (empty list = no exams)
"""

from __future__ import annotations

from collections import defaultdict
from enum import IntEnum

# (window_size, window_weight) pairs — applied in order
_PASSES   = [(2, 3), (3, 2), (4, 1)]
_SESSIONS = ["AM", "PM"]


class ColourWeight(IntEnum):
    RED    = 3
    YELLOW = 2
    GREEN  = 1
    UNSET  = 1


def _colour_weight(paper, registry) -> int:
    """Return the ColourWeight int for a paper via the registry difficulty lookup."""
    diff = registry.get_difficulty(paper.subject, paper.grade)
    if diff == "red":
        return int(ColourWeight.RED)
    if diff == "yellow":
        return int(ColourWeight.YELLOW)
    return int(ColourWeight.GREEN)


def _max_window_coeff(max_day: int) -> int:
    """Return the maximum per-day window coefficient across all days.

    For day *d* in a schedule spanning [0, max_day], the coefficient is the
    number of (weighted) convolution windows that contain *d*.  The hottest
    day is near the centre of the range.
    """
    if max_day < 0:
        return 1
    best = 0
    for d in range(max_day + 1):
        total = 0
        for w_size, w_weight in _PASSES:
            upper = max_day - w_size + 1
            if upper < 0:
                continue
            lo = max(0, d - w_size + 1)
            hi = min(d, upper)
            count = hi - lo + 1 if hi >= lo else 0
            total += w_weight * count
        if total > best:
            best = total
    return best if best else 1


class StudentStressCost:
    """
    Per-student overlap stress cost.

    Only penalises when a specific student has 2+ exams in the same
    convolution window.  For each such student in each such window the
    contribution is ``window_weight * sum(colour_weight of their exams)``.

    Normalised to [0, weight] by dividing by the theoretical worst case
    (all papers stacked on the hottest day).
    """

    def __init__(self, exam_tree, registry, weight: int) -> None:
        self._registry = registry
        self._weight   = weight
        # Normalisation base: sum over all students of the total colour
        # weight of every paper they sit.  Worst case = all exams on
        # the same day so every student's full colour-weight contributes.
        student_cw: dict[int, int] = defaultdict(int)
        for p in registry.all_papers():
            cw = _colour_weight(p, registry)
            for sid in p.student_ids:
                student_cw[sid] += cw
        self._total_student_base = sum(student_cw.values())

    # ── helpers ───────────────────────────────────────────────────────

    @staticmethod
    def _build_student_exams(schedule: dict, registry) -> dict[int, list[tuple[int, int]]]:
        """Return {student_id: [(day, colour_weight), ...]} from a schedule."""
        student_exams: dict[int, list[tuple[int, int]]] = defaultdict(list)
        for (day, _session), papers in schedule.items():
            for paper in papers:
                cw = _colour_weight(paper, registry)
                for sid in paper.student_ids:
                    student_exams[sid].append((day, cw))
        return student_exams

    # ── public API ────────────────────────────────────────────────────

    def compute(self, schedule: dict) -> float:
        if not schedule:
            return 0.0
        max_day = max(d for d, _ in schedule)
        student_exams = self._build_student_exams(schedule, self._registry)

        total = 0.0
        for sid, exams in student_exams.items():
            for w_size, w_weight in _PASSES:
                upper = max_day - w_size + 1
                if upper < 0:
                    continue
                for ws in range(upper + 1):
                    in_window = [cw for d, cw in exams if ws <= d < ws + w_size]
                    if len(in_window) >= 2:
                        total += w_weight * sum(in_window)

        max_raw = self._total_student_base * _max_window_coeff(max_day)
        if max_raw == 0:
            return 0.0
        return (total / max_raw) * self._weight

    def compute_windows(self, schedule: dict,
                        windows: list[tuple[int, int, int]]) -> float:
        """Compute cost contribution for a specific set of windows only."""
        if not schedule:
            return 0.0
        max_day = max(d for d, _ in schedule)
        student_exams = self._build_student_exams(schedule, self._registry)

        total = 0.0
        for ws, w_size, w_weight in windows:
            for sid, exams in student_exams.items():
                in_window = [cw for d, cw in exams if ws <= d < ws + w_size]
                if len(in_window) >= 2:
                    total += w_weight * sum(in_window)

        max_raw = self._total_student_base * _max_window_coeff(max_day)
        if max_raw == 0:
            return 0.0
        return (total / max_raw) * self._weight


class TeacherMarkingCost:
    """
    Convolution-based teacher marking load cost.

    Same convolution structure as StudentStressCost but:
      - Unit per slot: students that this teacher marks for the paper
      - No colour multiplier
      - Summed across all teachers in each window
    """

    def __init__(self, exam_tree, paper_registry, weight: int) -> None:
        self._weight      = weight
        # paper_label -> {teacher_code -> student_count}
        self._teacher_map: dict[str, dict[str, int]] = {}
        self._build_teacher_map(exam_tree, paper_registry)
        self._total_base = sum(
            sum(teachers.values())
            for teachers in self._teacher_map.values()
        )

    def _build_teacher_map(self, exam_tree, paper_registry) -> None:
        for paper in paper_registry.all_papers():
            grade_num  = paper.grade.replace("Gr", "").replace("gr", "")
            subj_label = f"{paper.subject}_{grade_num}"
            grade_node = exam_tree.grades.get(paper.grade)
            if not grade_node:
                continue
            exam_subj = grade_node.exam_subjects.get(subj_label)
            if not exam_subj:
                continue
            teacher_counts: dict[str, int] = {}
            for class_label, class_list in exam_subj.class_lists.items():
                parts = class_label.split("_")
                if len(parts) >= 3:
                    teacher = "_".join(parts[1:-1])
                    n = len(class_list.student_list.students)
                    teacher_counts[teacher] = teacher_counts.get(teacher, 0) + n
            if teacher_counts:
                self._teacher_map[paper.label] = teacher_counts

    def compute(self, schedule: dict) -> float:
        if not schedule:
            return 0.0
        max_day = max(d for d, _ in schedule)
        total_days = len({day for day, _ in schedule})
        total   = 0.0
        for w_size, w_weight in _PASSES:
            upper = max_day - w_size + 1
            if upper < 0:
                continue
            pass_total = 0.0
            for d in range(upper + 1):
                position_weight = (d + 1) / total_days
                teacher_totals: dict[str, int] = {}
                for day in range(d, d + w_size):
                    for session in _SESSIONS:
                        for paper in schedule.get((day, session), []):
                            for teacher, count in self._teacher_map.get(paper.label, {}).items():
                                teacher_totals[teacher] = teacher_totals.get(teacher, 0) + count
                pass_total += position_weight * sum(teacher_totals.values())
            total += pass_total * w_weight
        max_raw = self._total_base * _max_window_coeff(max_day)
        if max_raw == 0:
            return 0.0
        return (total / max_raw) * self._weight

    def compute_windows(self, schedule: dict,
                        windows: list[tuple[int, int, int]]) -> float:
        """Compute cost contribution for a specific set of windows only."""
        if not schedule:
            return 0.0
        max_day = max(d for d, _ in schedule)
        total_days = len({day for day, _ in schedule})
        total = 0.0
        for d, w_size, w_weight in windows:
            position_weight = (d + 1) / total_days
            teacher_totals: dict[str, int] = {}
            for day in range(d, d + w_size):
                for session in _SESSIONS:
                    for paper in schedule.get((day, session), []):
                        for teacher, count in self._teacher_map.get(paper.label, {}).items():
                            teacher_totals[teacher] = teacher_totals.get(teacher, 0) + count
            total += position_weight * sum(teacher_totals.values()) * w_weight
        max_raw = self._total_base * _max_window_coeff(max_day)
        if max_raw == 0:
            return 0.0
        return (total / max_raw) * self._weight


class TotalCost:
    def __init__(self, stress: StudentStressCost, marking: TeacherMarkingCost) -> None:
        self._stress  = stress
        self._marking = marking

    def compute(self, schedule: dict) -> float:
        return self._stress.compute(schedule) + self._marking.compute(schedule)

    def compute_windows(self, schedule: dict,
                        windows: list[tuple[int, int, int]]) -> float:
        return (self._stress.compute_windows(schedule, windows)
                + self._marking.compute_windows(schedule, windows))
