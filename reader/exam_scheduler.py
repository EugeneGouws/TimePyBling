"""
exam_scheduler.py
-----------------
Priority-based exam scheduler operating across ALL grades simultaneously.

Algorithm
---------
Step 1  Collect priority papers (MA and PH), sorted Gr12→Gr08 then P1→P3.
        Assign each to the slot that:
          a) does not clash with any already-assigned paper, AND
          b) maximises the minimum gap (in slots) to other papers with the
             same subject code within the SAME grade (e.g. MA_P1_Gr12 vs
             MA_P2_Gr12).  Papers from other grades have no spacing requirement.

Step 2  Collect remaining papers, sorted by student count desc, then grade desc.
        Assign with DSatur respecting the full clash graph and treating
        already-assigned slots as fixed.

Step 3  Spacing penalty pass.  Same-subject, same-grade papers in the same
        Mon–Fri week are flagged as warnings (per-grade diagnostic).
        Non-priority papers are swapped to a better slot if one exists with
        no new clashes introduced.

Output  ScheduleResult — list of ScheduledPaper (one per ExamPaper).

Configuration
-------------
  Sessions per day : 2  (AM, PM)
  Slot 0  = Day 1 AM,  Slot 1  = Day 1 PM,  Slot 2 = Day 2 AM, …
  Exam days: Monday–Friday only; weekends are skipped.
  Default start: next Thursday from the date the module is imported.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

from reader.exam_paper import ExamPaper, ExamPaperRegistry
from reader.exam_clash import build_paper_clash_graph


# ── Configuration ────────────────────────────────────────────────────────────

SESSIONS          = ["AM", "PM"]
EXAM_WEEKDAYS     = {0, 1, 2, 3, 4}          # Mon=0 … Fri=4
PRIORITY_SUBJECTS = {"MA", "PH"}


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


@dataclass
class ScheduleResult:
    scheduled   : list[ScheduledPaper]
    total_slots : int
    total_days  : int
    exact       : bool = True   # False if DSatur fallback was used somewhere


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


def _slot_to_date_session(slot: int, exam_days: list[date]) -> tuple[date, str]:
    day_idx  = slot // 2
    sess_idx = slot % 2
    return exam_days[day_idx], SESSIONS[sess_idx]


def _week_number(d: date) -> int:
    """ISO week number — used to detect same-week clashes."""
    return d.isocalendar()[1]


def _grade_sort_key(grade: str) -> int:
    """Gr12 → 12, Gr08 → 8.  Higher = scheduled first."""
    try:
        return int(grade.replace("Gr", "").replace("gr", ""))
    except ValueError:
        return 0


# ── Core scheduler ────────────────────────────────────────────────────────────

def build_schedule(
    registry        : ExamPaperRegistry,
    total_days      : int  = DEFAULT_TOTAL_DAYS,
    start_date      : date = DEFAULT_START_DATE,
    sessions_per_day: int  = 2,
) -> ScheduleResult:
    """
    Schedule all papers in the registry into a single cross-grade timetable.

    Parameters
    ----------
    registry         : ExamPaperRegistry  (all papers, all grades)
    total_days       : int  — total exam days available
    start_date       : date — first exam day
    sessions_per_day : int  — always 2 (AM / PM)

    Returns
    -------
    ScheduleResult
    """
    papers    = registry.all_papers()
    graph     = build_paper_clash_graph(papers)
    num_slots = total_days * sessions_per_day
    exam_days = _exam_dates(start_date, total_days)

    # assignment: label -> slot index (0-based)
    assignment: dict[str, int] = {}
    # which labels are priority
    priority_labels: set[str] = set()

    # ── Step 1: Priority placement ────────────────────────────────────────

    priority_papers = [
        p for p in papers if p.subject in PRIORITY_SUBJECTS
    ]
    priority_papers.sort(
        key=lambda p: (-_grade_sort_key(p.grade), p.paper_number)
    )

    for paper in priority_papers:
        forbidden = {assignment[nb] for nb in graph[paper.label]
                     if nb in assignment}

        # Find all valid slots
        valid_slots = [s for s in range(num_slots) if s not in forbidden]
        if not valid_slots:
            # No valid slot — fall back to least-clashing slot
            valid_slots = list(range(num_slots))

        # Among valid, pick the one that maximises min-gap to same subject+grade
        same_subj_slots = [
            assignment[lbl]
            for lbl in assignment
            if _label_to_parts(lbl)[0] == paper.subject
            and _label_to_parts(lbl)[2] == paper.grade
        ]

        def gap_score(s: int) -> int:
            if not same_subj_slots:
                return num_slots   # no existing same-subject — any slot is fine
            return min(abs(s - other) for other in same_subj_slots)

        best = max(valid_slots, key=gap_score)
        assignment[paper.label] = best
        priority_labels.add(paper.label)

    # ── Step 2: DSatur for remaining papers ───────────────────────────────

    remaining = [p for p in papers if p.label not in assignment]
    remaining.sort(
        key=lambda p: (-p.student_count(), -_grade_sort_key(p.grade))
    )

    # Used-slots per node = {slot: count} — for saturation heuristic
    saturation: dict[str, set[int]] = {p.label: set() for p in remaining}

    # Seed saturation from already-assigned neighbours
    for paper in remaining:
        for nb in graph[paper.label]:
            if nb in assignment:
                saturation[paper.label].add(assignment[nb])

    unassigned = {p.label: p for p in remaining}

    while unassigned:
        # Pick highest saturation, break ties by student count
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

        slot = 0
        while slot in forbidden and slot < num_slots:
            slot += 1
        # If we ran out of slots, pack into last slot (overflow warning later)
        if slot >= num_slots:
            slot = num_slots - 1

        assignment[chosen_label] = slot

        # Update saturation of remaining unassigned neighbours
        for nb in graph[chosen_label]:
            if nb in unassigned and nb != chosen_label:
                saturation[nb].add(slot)

        del unassigned[chosen_label]

    # ── Step 3: Spacing penalty pass ─────────────────────────────────────

    warnings: dict[str, list[str]] = {p.label: [] for p in papers}

    # Build slot→date map for all assigned slots
    def slot_date(s: int) -> date:
        return exam_days[s // 2] if (s // 2) < len(exam_days) else exam_days[-1]

    # Group by (subject, grade) — spacing diagnostic is per-grade
    from collections import defaultdict
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
                    msg = (f"{subj} {grade}: {la} (slot {sa+1}) and {lb} "
                           f"(slot {sb+1}) in same week "
                           f"({da.strftime('%d %b')})")
                    warnings[la].append(msg)
                    warnings[lb].append(msg)

                    # Try to swap non-priority paper to a better slot
                    for lbl in (la, lb):
                        if lbl in priority_labels or lbl in swapped:
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

    # ── Assemble result ───────────────────────────────────────────────────

    paper_map = {p.label: p for p in papers}
    scheduled: list[ScheduledPaper] = []

    for label, slot in sorted(assignment.items(), key=lambda x: x[1]):
        d, session = _slot_to_date_session(slot, exam_days)
        scheduled.append(ScheduledPaper(
            paper      = paper_map[label],
            slot_index = slot,
            date       = d,
            session    = session,
            warnings   = warnings.get(label, []),
        ))

    return ScheduleResult(
        scheduled   = scheduled,
        total_slots = num_slots,
        total_days  = total_days,
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
        w = "  ⚠" if sp.warnings else ""
        print(f"  Slot {sp.slot_index+1:>2}  {sp.date.strftime('%a %d %b')}  "
              f"{sp.session:<3}  {sp.paper.label:<18}  "
              f"{sp.paper.student_count():>3} students{w}")
