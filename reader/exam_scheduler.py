"""
exam_scheduler.py
-----------------
Three-phase constructive exam scheduler operating across ALL grades.

Algorithm
---------
Step 0    Pinned papers (including ST study papers) placed first at their
          pinned_slot.  ST papers reserve their slots for their grade.

Phase 1   RED + LINKED subjects — AM slots only, 5-day spacing between
          same-grade papers.  Linked partners auto-placed in PM of the
          same day.  Linked subjects are auto-promoted to red.
          Grades processed Gr12 → Gr08, most constrained first.

Phase 2   YELLOW subjects — fill gaps between red papers.  AM slots first;
          PM slots unlock when AM is exhausted.  Each paper placed in the
          slot that minimises incremental per-student overlap cost.
          Grades processed Gr12 → Gr08, most constrained first.

Phase 3   GREEN subjects — same logic as Phase 2.

Post      Swap-based hill-climb to locally minimise student stress cost.

Output    ScheduleResult — list of ScheduledPaper (one per ExamPaper).

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
from core.cost_function import _colour_weight, _PASSES, _max_window_coeff


# ── Configuration ────────────────────────────────────────────────────────────

SESSIONS      = ["AM", "PM"]
EXAM_WEEKDAYS = {0, 1, 2, 3, 4}          # Mon=0 … Fri=4


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
    pinned     : bool = False


@dataclass
class PenaltyEntry:
    constraint : str        # "day_density", "week_density", "teacher_load"
    papers     : list[str]  # paper labels involved
    entity     : str        # student ID string, teacher code, or ""
    value      : int        # penalty value


@dataclass
class ScheduleResult:
    scheduled          : list[ScheduledPaper]
    total_slots        : int
    total_days         : int
    exact              : bool = True
    sessions           : list[tuple[date, str]] = field(default_factory=list)
    student_cost       : float = 0.0
    penalty_log        : list[PenaltyEntry]      = field(default_factory=list)


# ── Retained helpers ─────────────────────────────────────────────────────────

def _exam_dates(start: date, count: int) -> list[date]:
    """Return `count` weekday dates starting from `start` (skipping weekends)."""
    dates: list[date] = []
    d = start
    while len(dates) < count:
        if d.weekday() in EXAM_WEEKDAYS:
            dates.append(d)
        d += timedelta(days=1)
    return dates


def _grade_sort_key(grade: str) -> int:
    """Gr12 → 12, Gr08 → 8.  Higher = scheduled first."""
    try:
        return int(grade.replace("Gr", "").replace("gr", ""))
    except ValueError:
        return 0



# ── New placement helpers ────────────────────────────────────────────────────

def _day_of_slot(slot: int) -> int:
    """Slot index to day index.  Slot 0,1 → day 0; slot 2,3 → day 1."""
    return slot // 2


def _is_am(slot: int) -> bool:
    """True if the slot is an AM slot (even index)."""
    return slot % 2 == 0


def _most_constrained_order(
    papers: list[ExamPaper],
    graph: dict[str, set[str]],
) -> list[ExamPaper]:
    """Sort papers by descending clash-graph degree, then descending student count."""
    return sorted(
        papers,
        key=lambda p: (-len(graph.get(p.label, set())), -p.student_count()),
    )


def _incremental_student_cost(
    paper: ExamPaper,
    candidate_slot: int,
    student_day_cws: dict[int, dict[int, list[int]]],
    registry: ExamPaperRegistry,
    max_day: int,
) -> int:
    """Marginal increase in per-student overlap cost from placing paper here.

    For each student in paper.student_ids, for each convolution window
    containing the candidate day, compute the cost delta from adding one
    more exam.  Returns 0 when no student gains an overlap.
    """
    candidate_day = _day_of_slot(candidate_slot)
    cw_new = _colour_weight(paper, registry)
    delta = 0

    for sid in paper.student_ids:
        day_cws = student_day_cws.get(sid, {})
        for w_size, w_weight in _PASSES:
            upper = max_day - w_size + 1
            if upper < 0:
                continue
            lo = max(0, candidate_day - w_size + 1)
            hi = min(candidate_day, upper)
            for ws in range(lo, hi + 1):
                existing_count = 0
                existing_cw_sum = 0
                for d in range(ws, ws + w_size):
                    if d == candidate_day:
                        continue  # don't count the paper being placed
                    for cw in day_cws.get(d, []):
                        existing_count += 1
                        existing_cw_sum += cw
                # Also count existing exams on candidate_day itself
                for cw in day_cws.get(candidate_day, []):
                    existing_count += 1
                    existing_cw_sum += cw

                if existing_count == 0:
                    pass  # 0 existing → new total 1 → no penalty
                elif existing_count == 1:
                    # 1 existing → new total 2 → new overlap created
                    delta += w_weight * (existing_cw_sum + cw_new)
                else:
                    # 2+ existing → adding one more → marginal cost
                    delta += w_weight * cw_new

    return delta


def _find_best_slot(
    paper: ExamPaper,
    available_slots: list[int],
    graph: dict[str, set[str]],
    assignment: dict[str, int],
    student_day_cws: dict[int, dict[int, list[int]]],
    registry: ExamPaperRegistry,
    max_day: int,
) -> int | None:
    """Pick the slot from available_slots that minimises incremental cost.

    Filters out clash-graph forbidden slots.
    Tiebreak: prefer AM (even index), then earliest slot.
    Returns None if no valid slot exists.
    """
    forbidden = {assignment[nb] for nb in graph.get(paper.label, set())
                 if nb in assignment}
    candidates = [s for s in available_slots if s not in forbidden]
    if not candidates:
        return None

    best_slot = None
    best_key = None
    for s in candidates:
        cost = _incremental_student_cost(paper, s, student_day_cws, registry, max_day)
        key = (cost, 0 if _is_am(s) else 1, s)
        if best_key is None or key < best_key:
            best_key = key
            best_slot = s

    return best_slot


def _get_available_slots(
    paper: ExamPaper,
    graph: dict[str, set[str]],
    assignment: dict[str, int],
    grade_reserved: dict[str, set[int]],
    num_slots: int,
    pm_unlocked: bool,
) -> tuple[list[int], bool]:
    """Return (available_slots, pm_unlocked).

    AM slots only while pm_unlocked is False.
    Once AM slots are exhausted, PM is unlocked permanently.
    """
    reserved = grade_reserved.get(paper.grade, set())
    forbidden = {assignment[nb] for nb in graph.get(paper.label, set())
                 if nb in assignment}

    if not pm_unlocked:
        am_slots = [s for s in range(0, num_slots, 2)
                    if s not in forbidden and s not in reserved]
        if am_slots:
            return am_slots, False
        pm_unlocked = True

    all_slots = [s for s in range(num_slots)
                 if s not in forbidden and s not in reserved]
    return all_slots, pm_unlocked


def _place_paper(
    paper: ExamPaper,
    slot: int,
    assignment: dict[str, int],
    student_day_cws: dict[int, dict[int, list[int]]],
    registry: ExamPaperRegistry,
) -> None:
    """Record a paper placement in assignment and student tracking structures."""
    assignment[paper.label] = slot
    day = _day_of_slot(slot)
    cw = _colour_weight(paper, registry)
    for sid in paper.student_ids:
        student_day_cws[sid][day].append(cw)


def _unplace_paper(
    paper: ExamPaper,
    slot: int,
    assignment: dict[str, int],
    student_day_cws: dict[int, dict[int, list[int]]],
    registry: ExamPaperRegistry,
) -> None:
    """Remove a paper placement from assignment and student tracking."""
    del assignment[paper.label]
    day = _day_of_slot(slot)
    cw = _colour_weight(paper, registry)
    for sid in paper.student_ids:
        cws = student_day_cws[sid][day]
        cws.remove(cw)
        if not cws:
            del student_day_cws[sid][day]


def _compute_total_student_cost(
    student_day_cws: dict[int, dict[int, list[int]]],
    max_day: int,
) -> float:
    """Compute the full per-student overlap cost across all convolution windows."""
    total = 0.0
    for sid, day_cws in student_day_cws.items():
        for w_size, w_weight in _PASSES:
            upper = max_day - w_size + 1
            if upper < 0:
                continue
            for ws in range(upper + 1):
                cw_sum = 0
                count = 0
                for d in range(ws, ws + w_size):
                    for cw in day_cws.get(d, []):
                        count += 1
                        cw_sum += cw
                if count >= 2:
                    total += w_weight * cw_sum
    return total


# ── Core scheduler ────────────────────────────────────────────────────────────

def build_schedule(
    registry         : ExamPaperRegistry,
    total_days       : int  = DEFAULT_TOTAL_DAYS,
    start_date       : date = DEFAULT_START_DATE,
    sessions_per_day : int  = 2,
    sessions         : list[tuple[date, str]] | None = None,
    exam_tree        = None,
    config           = None,
) -> ScheduleResult:
    """
    Schedule all papers using a three-phase constructive algorithm.

    Phase 1: RED + LINKED papers (AM only, 5-day spacing)
    Phase 2: YELLOW papers (fill gaps, AM first, PM when exhausted)
    Phase 3: GREEN papers (same as Phase 2)
    Post:    Swap-based hill-climb for local optimisation
    """
    papers = registry.all_papers()
    graph  = build_paper_clash_graph(papers)
    paper_map: dict[str, ExamPaper] = {p.label: p for p in papers}

    # ── Build slot list ──────────────────────────────────────────────────────

    if sessions is not None:
        exam_slots  = list(sessions)
        num_slots   = len(exam_slots)
        _total_days = len({d for d, _ in exam_slots})
    else:
        exam_days   = _exam_dates(start_date, total_days)
        exam_slots  = [(d, s) for d in exam_days for s in SESSIONS]
        num_slots   = len(exam_slots)
        _total_days = total_days

    max_day = (num_slots - 1) // 2 if num_slots > 0 else 0

    # ── Tracking structures ──────────────────────────────────────────────────

    assignment: dict[str, int] = {}
    pinned_label_set: set[str] = set()

    # student_day_cws[sid][day] = list of colour_weights of exams on that day
    student_day_cws: dict[int, dict[int, list[int]]] = defaultdict(
        lambda: defaultdict(list)
    )

    # ── Step 0: Pinned papers (including ST) ─────────────────────────────────

    pinned_papers   = [p for p in papers if p.pinned_slot is not None]
    unpinned_papers = [p for p in papers if p.pinned_slot is None]

    for p in pinned_papers:
        slot = min(p.pinned_slot, num_slots - 1)
        _place_paper(p, slot, assignment, student_day_cws, registry)
        pinned_label_set.add(p.label)

    # Grade-reserved slots from ST papers
    grade_reserved_slots: dict[str, set[int]] = defaultdict(set)
    for p in pinned_papers:
        if p.subject == "ST":
            grade_reserved_slots[p.grade].add(assignment[p.label])

    # ── Partition unpinned papers by colour ──────────────────────────────────

    # Identify linked labels — auto-promote to red
    linked_labels: set[str] = set()
    for p in papers:
        if p.links:
            linked_labels.add(p.label)
            for lnk in p.links:
                linked_labels.add(lnk)

    def _effective_colour(p: ExamPaper) -> str:
        if p.label in linked_labels:
            return "red"
        return registry.get_difficulty(p.subject, p.grade)

    grades_desc = sorted(
        {p.grade for p in unpinned_papers},
        key=_grade_sort_key,
        reverse=True,
    )

    pm_unlocked = False

    # ── PHASE 1: RED + LINKED (AM only, 5-day spacing) ──────────────────────

    for grade in grades_desc:
        red_papers = [
            p for p in unpinned_papers
            if p.grade == grade
            and _effective_colour(p) == "red"
            and p.label not in assignment
        ]
        red_papers = _most_constrained_order(red_papers, graph)

        for paper in red_papers:
            if paper.label in assignment:
                continue  # already placed as a linked partner

            # Determine if this paper has a linked partner
            partner = None
            for lnk in paper.links:
                if lnk in paper_map and lnk not in assignment:
                    partner = paper_map[lnk]
                    break

            # Available AM slots, not forbidden, not reserved
            reserved = grade_reserved_slots.get(paper.grade, set())
            forbidden = {assignment[nb] for nb in graph.get(paper.label, set())
                         if nb in assignment}
            am_slots = [s for s in range(0, num_slots, 2)
                        if s not in forbidden and s not in reserved]

            # Enforce 5-day minimum gap from other same-grade assigned slots
            same_grade_days = {
                _day_of_slot(assignment[lbl])
                for lbl in assignment
                if paper_map[lbl].grade == grade
            }
            filtered = [s for s in am_slots
                        if all(abs(_day_of_slot(s) - d) >= 5
                               for d in same_grade_days)]
            if not filtered:
                filtered = am_slots  # relax spacing if impossible

            if not filtered:
                # No AM slots at all — use any available slot
                filtered = [s for s in range(num_slots)
                            if s not in forbidden and s not in reserved]

            best = _find_best_slot(paper, filtered, graph, assignment,
                                    student_day_cws, registry, max_day)
            if best is None:
                best = filtered[0] if filtered else 0
            _place_paper(paper, best, assignment, student_day_cws, registry)

            # Place linked partner in PM of same day
            if partner is not None and partner.label not in assignment:
                pm_slot = best + 1 if _is_am(best) else best - 1
                if not _is_am(pm_slot):
                    # pm_slot is indeed PM
                    forbidden_p = {assignment[nb]
                                   for nb in graph.get(partner.label, set())
                                   if nb in assignment}
                    reserved_p = grade_reserved_slots.get(partner.grade, set())
                    if (0 <= pm_slot < num_slots
                            and pm_slot not in forbidden_p
                            and pm_slot not in reserved_p):
                        _place_paper(partner, pm_slot, assignment,
                                      student_day_cws, registry)
                    else:
                        # Fallback: best available AM slot for partner
                        p_slots = [s for s in range(0, num_slots, 2)
                                   if s not in forbidden_p
                                   and s not in reserved_p]
                        if not p_slots:
                            p_slots = [s for s in range(num_slots)
                                       if s not in forbidden_p
                                       and s not in reserved_p] \
                                      or list(range(num_slots))
                        fb = _find_best_slot(partner, p_slots, graph,
                                              assignment, student_day_cws,
                                              registry, max_day)
                        if fb is None:
                            fb = p_slots[0]
                        _place_paper(partner, fb, assignment,
                                      student_day_cws, registry)

    # ── PHASE 2: YELLOW (fill gaps, AM first then PM) ────────────────────────

    for grade in grades_desc:
        yellow_papers = [
            p for p in unpinned_papers
            if p.grade == grade
            and _effective_colour(p) == "yellow"
            and p.label not in assignment
        ]
        yellow_papers = _most_constrained_order(yellow_papers, graph)

        for paper in yellow_papers:
            available, pm_unlocked = _get_available_slots(
                paper, graph, assignment, grade_reserved_slots,
                num_slots, pm_unlocked,
            )
            if not available:
                available = list(range(num_slots))
            best = _find_best_slot(paper, available, graph, assignment,
                                    student_day_cws, registry, max_day)
            if best is None:
                best = available[0]
            _place_paper(paper, best, assignment, student_day_cws, registry)

    # ── PHASE 3: GREEN (same logic as Phase 2) ──────────────────────────────

    for grade in grades_desc:
        green_papers = [
            p for p in unpinned_papers
            if p.grade == grade
            and _effective_colour(p) == "green"
            and p.label not in assignment
        ]
        green_papers = _most_constrained_order(green_papers, graph)

        for paper in green_papers:
            available, pm_unlocked = _get_available_slots(
                paper, graph, assignment, grade_reserved_slots,
                num_slots, pm_unlocked,
            )
            if not available:
                available = list(range(num_slots))
            best = _find_best_slot(paper, available, graph, assignment,
                                    student_day_cws, registry, max_day)
            if best is None:
                best = available[0]
            _place_paper(paper, best, assignment, student_day_cws, registry)

    # ── POST-PLACEMENT HILL-CLIMB ────────────────────────────────────────────
    # Simple paper-move search: for each non-pinned paper, try every valid
    # slot; accept if incremental cost is lower than current position.

    non_pinned = [lbl for lbl in assignment if lbl not in pinned_label_set]

    for _pass in range(20):
        improved = False
        for lbl in non_pinned:
            paper = paper_map[lbl]
            old_slot = assignment[lbl]
            reserved = grade_reserved_slots.get(paper.grade, set())
            forbidden = {assignment[nb] for nb in graph.get(lbl, set())
                         if nb in assignment and nb != lbl}

            # Cost of current position
            _unplace_paper(paper, old_slot, assignment, student_day_cws, registry)
            old_cost = _incremental_student_cost(
                paper, old_slot, student_day_cws, registry, max_day)

            best_cost = old_cost
            best_slot = old_slot

            for s in range(num_slots):
                if s == old_slot or s in forbidden or s in reserved:
                    continue
                cost = _incremental_student_cost(
                    paper, s, student_day_cws, registry, max_day)
                key_new = (cost, 0 if _is_am(s) else 1, s)
                key_best = (best_cost, 0 if _is_am(best_slot) else 1, best_slot)
                if key_new < key_best:
                    best_cost = cost
                    best_slot = s

            _place_paper(paper, best_slot, assignment, student_day_cws, registry)
            if best_slot != old_slot:
                improved = True

        if not improved:
            break

    # ── Compute final student cost ───────────────────────────────────────────

    student_cost = _compute_total_student_cost(student_day_cws, max_day)

    # ── Assemble result ──────────────────────────────────────────────────────

    scheduled: list[ScheduledPaper] = []
    for label, slot in sorted(assignment.items(), key=lambda x: x[1]):
        slot_clamped = min(slot, len(exam_slots) - 1)
        d, session = exam_slots[slot_clamped]
        scheduled.append(ScheduledPaper(
            paper      = paper_map[label],
            slot_index = slot,
            date       = d,
            session    = session,
            pinned     = label in pinned_label_set,
        ))

    return ScheduleResult(
        scheduled    = scheduled,
        total_slots  = num_slots,
        total_days   = _total_days,
        sessions     = exam_slots,
        student_cost = student_cost,
    )


# ── Standalone smoke test ────────────────────────────────────────────────────

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
    print(f"\n{len(result.scheduled)} papers scheduled over {result.total_days} days")
    print(f"Student cost: {result.student_cost:.1f}\n")
    for sp in result.scheduled:
        pin = " (pinned)" if sp.pinned else ""
        print(f"  Slot {sp.slot_index + 1:>2}  {sp.date.strftime('%a %d %b')}  "
              f"{sp.session:<3}  {sp.paper.label:<18}  "
              f"{sp.paper.student_count():>3} students{pin}")
