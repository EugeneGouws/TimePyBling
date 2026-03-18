"""
exam_scheduler.py
-----------------
Priority-based exam scheduler operating across ALL grades simultaneously.

Algorithm
---------
Step 0  Pinned papers are placed first at their pinned_slot.
        Pin-on-pin clashes are detected and surfaced as warnings but
        do not block placement — the user chose those slots deliberately.

Step 1  Collect priority papers (MA and PH), sorted Gr12→Gr08 then P1→P3.
        Assign each using _pick_spread_slot (even load + gap maximisation).

Step 2  Collect remaining papers via DSatur saturation ordering.
        Each paper is placed using _pick_spread_slot (not first-available).

Step 3  Spacing penalty pass — same-subject same-grade in the same ISO week
        produces a warning; non-priority papers are swapped to a better slot
        if one exists with no new clashes.

Step 4  Teacher marking load diagnostic.

Step 5  Student-load hill-climb.  Minimises total quadratic clustering cost:
          day_penalty(k)  = k*(k-1) * W_DAY   for k exams on the same day
          week_penalty(k) = (k-1)*(k+6) // 2  for k exams in the same ISO week
        For each non-pinned paper, try moving it to every other slot.
        Accept the move if it reduces total student cost and introduces no
        clash.  Repeat until no improvement (max 20 passes).
        Cost stored in ScheduleResult.student_cost.

Output  ScheduleResult — list of ScheduledPaper (one per ExamPaper).

Penalty constants
-----------------
  W_DAY  = 5   →  2 exams same day  → per-student cost 10
                   3 exams same day  → per-student cost 30
  week_penalty:    1 exam  → 0,  2 exams → 4,  3 exams → 9  (quadratic)

Configuration
-------------
  Sessions per day : 2  (AM, PM)
  Slot 0  = Day 1 AM,  Slot 1  = Day 1 PM,  Slot 2 = Day 2 AM, …
  Exam days: Monday–Friday only; weekends are skipped.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta

from reader.exam_paper import ExamPaper, ExamPaperRegistry
from reader.exam_clash import build_paper_clash_graph


# ── Configuration ────────────────────────────────────────────────────────────

SESSIONS          = ["AM", "PM"]
EXAM_WEEKDAYS     = {0, 1, 2, 3, 4}          # Mon=0 … Fri=4
PRIORITY_SUBJECTS = {"MA", "PH"}

# ── Student load penalty weights ──────────────────────────────────────────────
W_DAY = 5    # k exams on same day  → k*(k-1)*W_DAY per student  (k=2 → 10)
             # week_penalty uses fixed formula (k-1)*(k+6)//2    (k=2 → 4, k=3 → 9)


def _day_penalty(k: int) -> int:
    """Per-student cost for k exams on the same calendar day."""
    return k * (k - 1) * W_DAY if k > 1 else 0


def _week_penalty(k: int) -> int:
    """Per-student cost for k exams in the same ISO week.
    k=1→0, k=2→4, k=3→9, k=4→15, … (quadratic growth)."""
    return (k - 1) * (k + 6) // 2 if k > 1 else 0


def _next_thursday(from_date: date | None = None) -> date:
    d = from_date or date.today()
    days_until = (3 - d.weekday()) % 7 or 7   # 3 = Thursday
    return d + timedelta(days=days_until)


DEFAULT_START_DATE  = _next_thursday()
DEFAULT_TOTAL_DAYS  = 10


# ── Data classes ─────────────────────────────────────────────────────────────

@dataclass
class ScheduledPaper:
    paper      : ExamPaper
    slot_index : int
    date       : date
    session    : str           # "AM" or "PM"
    warnings   : list[str] = field(default_factory=list)
    pinned     : bool = False


@dataclass
class ScheduleResult:
    scheduled          : list[ScheduledPaper]
    total_slots        : int
    total_days         : int
    exact              : bool = True   # False if DSatur fallback was used somewhere
    sessions           : list[tuple[date, str]] = field(default_factory=list)
    pin_clash_warnings : dict[str, list[str]]   = field(default_factory=dict)
    teacher_warnings   : list[str]               = field(default_factory=list)
    student_cost       : int = 0


# ── Internal helpers ──────────────────────────────────────────────────────────

def _exam_dates(start: date, count: int) -> list[date]:
    """Return `count` weekday dates starting from `start` (skipping weekends)."""
    dates: list[date] = []
    d = start
    while len(dates) < count:
        if d.weekday() in EXAM_WEEKDAYS:
            dates.append(d)
        d += timedelta(days=1)
    return dates


def _week_number(d: date) -> int:
    """ISO week number — used to detect same-week clashes."""
    return d.isocalendar()[1]


def _grade_sort_key(grade: str) -> int:
    """Gr12 → 12, Gr08 → 8.  Higher = scheduled first."""
    try:
        return int(grade.replace("Gr", "").replace("gr", ""))
    except ValueError:
        return 0


def _pick_spread_slot(
    valid_slots : list[int],
    assignment  : dict[str, int],
    paper_label : str,
    num_slots   : int,
) -> int:
    """
    Choose the slot from valid_slots that best spreads the schedule:
      1. Fewest papers already in that slot  (primary — even load)
      2. Maximum minimum gap to same-subject same-grade papers  (secondary)
    """
    load: dict[int, int] = defaultdict(int)
    for s in assignment.values():
        load[s] += 1

    subj, _, grade = _label_to_parts(paper_label)
    same_sg = [
        assignment[lbl] for lbl in assignment
        if _label_to_parts(lbl)[0] == subj and _label_to_parts(lbl)[2] == grade
    ]

    def score(s: int) -> tuple:
        gap = min(abs(s - o) for o in same_sg) if same_sg else num_slots
        return (-load[s], gap)

    return max(valid_slots, key=score)


# ── Teacher code extraction ───────────────────────────────────────────────────

def _get_teacher_codes(exam_tree, subject: str, grade: str) -> set[str]:
    """
    Extract teacher codes for a subject+grade from the ExamTree.
    Returns empty set if exam_tree is None or the subject is not found.
    """
    if exam_tree is None:
        return set()
    grade_num  = grade.replace("Gr", "").replace("gr", "")
    subj_label = f"{subject}_{grade_num}"
    grade_node = exam_tree.grades.get(grade)
    if not grade_node:
        return set()
    exam_subj = grade_node.exam_subjects.get(subj_label)
    if not exam_subj:
        return set()
    teachers: set[str] = set()
    for class_label in exam_subj.class_lists:
        parts = class_label.split("_")
        if len(parts) >= 3:
            teachers.add("_".join(parts[1:-1]))
    return teachers


def _teacher_marking_warnings(
    assignment : dict[str, int],
    paper_map  : dict[str, ExamPaper],
    exam_tree,
) -> list[str]:
    """
    Scan for same subject + same teacher code appearing in multiple grades
    in the same slot.  Returns a list of human-readable warning strings.
    """
    slot_subject_papers: dict[tuple[int, str], list[ExamPaper]] = defaultdict(list)
    for label, slot in assignment.items():
        p = paper_map[label]
        slot_subject_papers[(slot, p.subject)].append(p)

    result: list[str] = []
    for (slot, subj), papers_in_slot in sorted(slot_subject_papers.items()):
        if len(papers_in_slot) < 2:
            continue
        teacher_to_grades: dict[str, list[str]] = defaultdict(list)
        for p in papers_in_slot:
            for teacher in _get_teacher_codes(exam_tree, p.subject, p.grade):
                teacher_to_grades[teacher].append(p.grade)
        for teacher, grades in sorted(teacher_to_grades.items()):
            if len(grades) >= 2:
                grade_str = " and ".join(sorted(grades))
                result.append(
                    f"Slot {slot + 1} — {subj}: {teacher} teaches {grade_str} "
                    f"simultaneously — marking load conflict"
                )
    return result


# ── Core scheduler ────────────────────────────────────────────────────────────

def build_schedule(
    registry         : ExamPaperRegistry,
    total_days       : int  = DEFAULT_TOTAL_DAYS,
    start_date       : date = DEFAULT_START_DATE,
    sessions_per_day : int  = 2,
    sessions         : list[tuple[date, str]] | None = None,
    exam_tree        = None,
    w_day            : int  = W_DAY,
    w_week           : int  = 1,
) -> ScheduleResult:
    """
    Schedule all papers in the registry into a single cross-grade timetable.

    Parameters
    ----------
    registry         : ExamPaperRegistry  (all papers, all grades)
    total_days       : int  — total exam days (ignored when sessions is provided)
    start_date       : date — first exam day  (ignored when sessions is provided)
    sessions_per_day : int  — always 2 (AM / PM)
    sessions         : explicit list of (date, session) tuples; when provided
                       the total_days / start_date params are ignored
    exam_tree        : ExamTree — required for teacher marking diagnostics;
                       pass None to skip that check

    Returns
    -------
    ScheduleResult
    """
    papers = registry.all_papers()
    graph  = build_paper_clash_graph(papers)

    # ── Build slot list ───────────────────────────────────────────────────────

    if sessions is not None:
        exam_slots  = list(sessions)
        num_slots   = len(exam_slots)
        _total_days = len({d for d, _ in exam_slots})
    else:
        exam_days   = _exam_dates(start_date, total_days)
        exam_slots  = [(d, s) for d in exam_days for s in SESSIONS]
        num_slots   = len(exam_slots)
        _total_days = total_days

    def slot_date(s: int) -> date:
        idx = min(s, len(exam_slots) - 1)
        return exam_slots[idx][0]

    assignment      : dict[str, int]   = {}
    priority_labels : set[str]         = set()
    pin_clash_warnings: dict[str, list[str]] = {}

    # ── Step 0: Pinned papers ──────────────────────────────────────────────────

    pinned_papers   = [p for p in papers if p.pinned_slot is not None]
    unpinned_papers = [p for p in papers if p.pinned_slot is None]
    pinned_label_set: set[str] = set()

    pinned_by_slot: dict[int, list[ExamPaper]] = defaultdict(list)
    for p in pinned_papers:
        slot = min(p.pinned_slot, num_slots - 1)
        assignment[p.label] = slot
        pinned_by_slot[slot].append(p)
        pinned_label_set.add(p.label)

    # Detect pin-on-pin clashes
    for slot, group in pinned_by_slot.items():
        for i, a in enumerate(group):
            for b in group[i + 1:]:
                if b.label in graph[a.label]:
                    msg = f"Clash with {b.label} (both pinned to slot {slot + 1})"
                    pin_clash_warnings.setdefault(a.label, []).append(msg)
                    pin_clash_warnings.setdefault(b.label, []).append(msg)

    # ── Step 1: Priority placement ─────────────────────────────────────────────

    priority_papers = [
        p for p in unpinned_papers if p.subject in PRIORITY_SUBJECTS
    ]
    priority_papers.sort(
        key=lambda p: (-_grade_sort_key(p.grade), p.paper_number)
    )

    for paper in priority_papers:
        forbidden = {assignment[nb] for nb in graph[paper.label]
                     if nb in assignment}
        valid_slots = [s for s in range(num_slots) if s not in forbidden] \
                      or list(range(num_slots))
        best = _pick_spread_slot(valid_slots, assignment, paper.label, num_slots)
        assignment[paper.label] = best
        priority_labels.add(paper.label)

    # ── Step 2: DSatur for remaining papers — spread across all slots ──────────

    remaining = [p for p in unpinned_papers if p.label not in assignment]
    remaining.sort(
        key=lambda p: (-p.student_count(), -_grade_sort_key(p.grade))
    )

    saturation: dict[str, set[int]] = {p.label: set() for p in remaining}
    for paper in remaining:
        for nb in graph[paper.label]:
            if nb in assignment:
                saturation[paper.label].add(assignment[nb])

    unassigned = {p.label: p for p in remaining}

    while unassigned:
        chosen_label = max(
            unassigned,
            key=lambda lbl: (
                len(saturation[lbl]),
                unassigned[lbl].student_count(),
                _grade_sort_key(unassigned[lbl].grade),
            )
        )
        forbidden = {assignment[nb] for nb in graph[chosen_label]
                     if nb in assignment}
        valid = [s for s in range(num_slots) if s not in forbidden] \
                or list(range(num_slots))
        slot = _pick_spread_slot(valid, assignment, chosen_label, num_slots)
        assignment[chosen_label] = slot

        for nb in graph[chosen_label]:
            if nb in unassigned and nb != chosen_label:
                saturation[nb].add(slot)

        del unassigned[chosen_label]

    # ── Step 3: Spacing penalty pass ───────────────────────────────────────────

    warnings: dict[str, list[str]] = {p.label: [] for p in papers}

    # Seed warnings with any pin clash messages
    for label, msgs in pin_clash_warnings.items():
        warnings.setdefault(label, []).extend(msgs)

    by_subject_grade: dict[tuple[str, str], list[str]] = defaultdict(list)
    for p in papers:
        by_subject_grade[(p.subject, p.grade)].append(p.label)

    swapped: set[str] = set()

    for (subj, grade), labels in by_subject_grade.items():
        assigned_labels = [lbl for lbl in labels if lbl in assignment]
        for i in range(len(assigned_labels)):
            for j in range(i + 1, len(assigned_labels)):
                la, lb = assigned_labels[i], assigned_labels[j]
                sa, sb = assignment[la], assignment[lb]
                da, db = slot_date(sa), slot_date(sb)
                if _week_number(da) == _week_number(db):
                    msg = (f"{subj} {grade}: {la} (slot {sa + 1}) and {lb} "
                           f"(slot {sb + 1}) in same week "
                           f"({da.strftime('%d %b')})")
                    warnings[la].append(msg)
                    warnings[lb].append(msg)

                    # Try to swap non-priority, non-pinned paper to a better slot
                    for lbl in (la, lb):
                        if (lbl in priority_labels or lbl in swapped
                                or lbl in pinned_label_set):
                            continue
                        forbidden = {assignment[nb] for nb in graph[lbl]
                                     if nb in assignment and nb != lbl}
                        same_subj_slots = [
                            assignment[other]
                            for other in assigned_labels
                            if other != lbl and other in assignment
                        ]
                        candidates = [
                            s for s in range(num_slots)
                            if s not in forbidden
                        ]
                        if not candidates:
                            continue
                        current_score = (
                            min(abs(assignment[lbl] - o) for o in same_subj_slots)
                            if same_subj_slots else 0
                        )
                        for s in candidates:
                            new_score = (
                                min(abs(s - o) for o in same_subj_slots)
                                if same_subj_slots else 0
                            )
                            new_d = slot_date(s)
                            other_slots = [
                                assignment[other]
                                for other in assigned_labels
                                if other != lbl and other in assignment
                            ]
                            same_week = any(
                                _week_number(new_d) == _week_number(slot_date(o))
                                for o in other_slots
                            )
                            if new_score > current_score and not same_week:
                                assignment[lbl] = s
                                swapped.add(lbl)
                                break

    # ── Step 4: Teacher marking load diagnostic ────────────────────────────────

    paper_map = {p.label: p for p in papers}
    teacher_warnings: list[str] = []
    if exam_tree is not None:
        teacher_warnings = _teacher_marking_warnings(assignment, paper_map, exam_tree)

    # ── Step 5: Student-load hill-climb ────────────────────────────────────────

    def _dp(k: int) -> int:
        return k * (k - 1) * w_day if k > 1 else 0

    def _wp(k: int) -> int:
        return ((k - 1) * (k + 6) // 2) * w_week if k > 1 else 0

    # Build per-student day/week exam counts
    s_day_c: dict[int, dict] = defaultdict(lambda: defaultdict(int))
    s_wk_c:  dict[int, dict] = defaultdict(lambda: defaultdict(int))
    for lbl, sl in assignment.items():
        d, _ = exam_slots[min(sl, len(exam_slots) - 1)]
        wk = _week_number(d)
        for sid in paper_map[lbl].student_ids:
            s_day_c[sid][d] += 1
            s_wk_c[sid][wk] += 1

    def _total_student_cost() -> int:
        total = 0
        for dc in s_day_c.values():
            for k in dc.values():
                total += _dp(k)
        for wc in s_wk_c.values():
            for k in wc.values():
                total += _wp(k)
        return total

    student_cost = _total_student_cost()
    non_pinned = [lbl for lbl in assignment if lbl not in pinned_label_set]

    for _pass in range(20):
        improved = False
        for lbl in non_pinned:
            sl_old = assignment[lbl]
            d_old, _ = exam_slots[min(sl_old, len(exam_slots) - 1)]
            wk_old = _week_number(d_old)
            forbidden = {assignment[nb] for nb in graph[lbl]
                         if nb in assignment and nb != lbl}

            best_delta = 0
            best_slot  = None

            for sl_new in range(num_slots):
                if sl_new == sl_old or sl_new in forbidden:
                    continue
                d_new, _ = exam_slots[min(sl_new, len(exam_slots) - 1)]
                wk_new = _week_number(d_new)

                delta = 0
                for sid in paper_map[lbl].student_ids:
                    if d_old != d_new:
                        k_do = s_day_c[sid].get(d_old, 0)
                        k_dn = s_day_c[sid].get(d_new, 0)
                        delta += (_dp(k_do - 1) - _dp(k_do)
                                  + _dp(k_dn + 1) - _dp(k_dn))
                    if wk_old != wk_new:
                        k_wo = s_wk_c[sid].get(wk_old, 0)
                        k_wn = s_wk_c[sid].get(wk_new, 0)
                        delta += (_wp(k_wo - 1) - _wp(k_wo)
                                  + _wp(k_wn + 1) - _wp(k_wn))

                if delta < best_delta:
                    best_delta = delta
                    best_slot  = sl_new

            if best_slot is not None:
                sl_new = best_slot
                d_new, _ = exam_slots[min(sl_new, len(exam_slots) - 1)]
                wk_new = _week_number(d_new)
                # Update counts
                for sid in paper_map[lbl].student_ids:
                    s_day_c[sid][d_old] -= 1
                    if s_day_c[sid][d_old] == 0:
                        del s_day_c[sid][d_old]
                    s_day_c[sid][d_new] += 1
                    s_wk_c[sid][wk_old] -= 1
                    if s_wk_c[sid][wk_old] == 0:
                        del s_wk_c[sid][wk_old]
                    s_wk_c[sid][wk_new] += 1
                assignment[lbl] = sl_new
                student_cost += best_delta
                improved = True

        if not improved:
            break

    # ── Assemble result ────────────────────────────────────────────────────────

    scheduled: list[ScheduledPaper] = []

    for label, slot in sorted(assignment.items(), key=lambda x: x[1]):
        slot_clamped = min(slot, len(exam_slots) - 1)
        d, session   = exam_slots[slot_clamped]
        scheduled.append(ScheduledPaper(
            paper      = paper_map[label],
            slot_index = slot,
            date       = d,
            session    = session,
            warnings   = warnings.get(label, []),
            pinned     = label in pinned_label_set,
        ))

    return ScheduleResult(
        scheduled          = scheduled,
        total_slots        = num_slots,
        total_days         = _total_days,
        sessions           = exam_slots,
        pin_clash_warnings = pin_clash_warnings,
        teacher_warnings   = teacher_warnings,
        student_cost       = student_cost,
    )


# ── Label parsing helper ──────────────────────────────────────────────────────

def _label_to_parts(label: str) -> tuple[str, str, str]:
    """
    "MA_P1_Gr12"  ->  ("MA", "P1", "Gr12")
    Returns ("", "", "") on unexpected format.
    """
    parts = label.split("_")
    if len(parts) >= 3:
        return parts[0], parts[1], parts[2]
    return "", "", ""


# ── Standalone smoke test ─────────────────────────────────────────────────────

if __name__ == "__main__":
    from reader.exam_paper import ExamPaper, ExamPaperRegistry

    mock_papers = [
        ExamPaper("Gr12", "MA", 1, student_ids={1, 2, 3, 4}),
        ExamPaper("Gr11", "MA", 1, student_ids={5, 6, 7, 8}),
        ExamPaper("Gr12", "EN", 1, student_ids={1, 2, 5, 6}),
        ExamPaper("Gr12", "PH", 1, student_ids={3, 4, 7, 8}),
        ExamPaper("Gr11", "EN", 1, student_ids={5, 6, 9, 10}),
    ]
    reg = ExamPaperRegistry()
    for p in mock_papers:
        reg._papers[p.label] = p

    result = build_schedule(reg, total_days=5, start_date=date(2026, 10, 5))
    print(f"\n{len(result.scheduled)} papers scheduled over {result.total_days} days\n")
    for sp in result.scheduled:
        pin = " 📌" if sp.pinned else ""
        w   = "  ⚠" if sp.warnings else ""
        print(f"  Slot {sp.slot_index + 1:>2}  {sp.date.strftime('%a %d %b')}  "
              f"{sp.session:<3}  {sp.paper.label:<18}  "
              f"{sp.paper.student_count():>3} students{pin}{w}")
