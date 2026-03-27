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


class StudentStressCost:
    """
    Convolution-based student stress cost.

    For each pass (window_size, window_weight):
        pass_total = sum over all windows of sum over all (grade, day, session):
            student_count(paper, grade) * colour_weight(paper)
        total += pass_total * window_weight
    Return total * (weight / 100).
    """

    def __init__(self, exam_tree, registry, weight: int) -> None:
        self._registry = registry
        self._weight   = weight
        self._grades   = list(exam_tree.grades.keys())

    def compute(self, schedule: dict) -> float:
        if not schedule:
            return 0.0
        max_day = max(d for d, _ in schedule)
        total   = 0.0
        for w_size, w_weight in _PASSES:
            upper = max_day - w_size + 1
            if upper < 0:
                continue
            pass_total = 0
            for d in range(upper + 1):
                for grade in self._grades:
                    window_score = 0
                    for day in range(d, d + w_size):
                        for session in _SESSIONS:
                            for paper in schedule.get((day, session), []):
                                if paper.grade == grade:
                                    n = paper.student_count()
                                    if n:
                                        window_score += n * _colour_weight(paper, self._registry)
                    pass_total += window_score
            total += pass_total * w_weight
        return total * (self._weight / 100)

    def compute_windows(self, schedule: dict,
                        windows: list[tuple[int, int, int]]) -> float:
        """Compute cost contribution for a specific set of windows only."""
        total = 0.0
        for d, w_size, w_weight in windows:
            for grade in self._grades:
                window_score = 0
                for day in range(d, d + w_size):
                    for session in _SESSIONS:
                        for paper in schedule.get((day, session), []):
                            if paper.grade == grade:
                                n = paper.student_count()
                                if n:
                                    window_score += n * _colour_weight(paper, self._registry)
                total += window_score * w_weight
        return total * (self._weight / 100)


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
        total   = 0.0
        for w_size, w_weight in _PASSES:
            upper = max_day - w_size + 1
            if upper < 0:
                continue
            pass_total = 0
            for d in range(upper + 1):
                teacher_totals: dict[str, int] = {}
                for day in range(d, d + w_size):
                    for session in _SESSIONS:
                        for paper in schedule.get((day, session), []):
                            for teacher, count in self._teacher_map.get(paper.label, {}).items():
                                teacher_totals[teacher] = teacher_totals.get(teacher, 0) + count
                pass_total += sum(teacher_totals.values())
            total += pass_total * w_weight
        return total * (self._weight / 100)

    def compute_windows(self, schedule: dict,
                        windows: list[tuple[int, int, int]]) -> float:
        """Compute cost contribution for a specific set of windows only."""
        total = 0.0
        for d, w_size, w_weight in windows:
            teacher_totals: dict[str, int] = {}
            for day in range(d, d + w_size):
                for session in _SESSIONS:
                    for paper in schedule.get((day, session), []):
                        for teacher, count in self._teacher_map.get(paper.label, {}).items():
                            teacher_totals[teacher] = teacher_totals.get(teacher, 0) + count
            total += sum(teacher_totals.values()) * w_weight
        return total * (self._weight / 100)


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
