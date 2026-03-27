"""
exam_scheduler.py
-----------------
Priority-based exam scheduler operating across ALL grades simultaneously.

Algorithm
---------
Step 0    Pinned papers (including ST study papers) are placed first at their
          pinned_slot.  Pin-on-pin clashes are detected and surfaced as
          warnings but do not block placement.  ST papers additionally
          reserve their slots for their grade (grade_reserved_slots).

Step 0.5  Linked paper pre-placement.  Pairs sorted by combined student
          count.  Primary at best slot; partner at slot±1 or fallback.
          Difficulty constraints are checked; linked partners are exempt
          from each other's difficulty rules.

Step 1    Collect priority papers (MA and PH), sorted Gr12→Gr08 then P1→P3.
          Assign each using _pick_spread_slot (even load + gap maximisation).
          Respects grade_reserved_slots and difficulty constraints.

Step 2    Collect remaining papers via DSatur saturation ordering.
          Each paper is placed using _pick_spread_slot (not first-available).
          Respects grade_reserved_slots and difficulty constraints.

Step 3    Spacing penalty pass — same-subject same-grade in the same ISO week
          produces a warning; non-priority papers are swapped to a better slot
          if one exists with no new clashes.

Step 4    Teacher marking load diagnostic.

Step 5    Student-load hill-climb.  Minimises total quadratic clustering cost:
            day_penalty(k)  = k*(k-1) * W_DAY   for k exams on the same day
            week_penalty(k) = (k-1)*(k+6) // 2  for k exams in the same ISO week
          For each non-pinned paper, try moving it to every other slot.
          Accept the move if it reduces total student cost and introduces no
          clash.  Respects grade_reserved_slots and difficulty constraints.
          AM-first tiebreak: equal-cost moves prefer AM slots (even index).
          Repeat until no improvement (max 20 passes).
          Cost stored in ScheduleResult.student_cost.

Output  ScheduleResult — list of ScheduledPaper (one per ExamPaper).

Penalty constants  (defaults; overridden by CostConfig when provided)
----------------------------------------------------------------------
  day_density_factor = 5   →  2 exams same day  → per-student cost 10
                               3 exams same day  → per-student cost 30
  week_density_base  = 6   →  k=2 → 4,  k=3 → 9  (quadratic)
  same_week_penalty  = 1   →  multiplier on week cost; 0 disables spacing swaps
  teacher_load_penalty = 1 →  reserved for penalty log (C4b)

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
    exact              : bool = True   # False if DSatur fallback was used somewhere
    sessions           : list[tuple[date, str]] = field(default_factory=list)
    pin_clash_warnings : dict[str, list[str]]   = field(default_factory=dict)
    teacher_warnings   : list[str]               = field(default_factory=list)
    student_cost       : int = 0
    penalty_log        : list[PenaltyEntry]      = field(default_factory=list)


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


# ── Difficulty feasibility check ──────────────────────────────────────────────

def _difficulty_allows(
    registry   : ExamPaperRegistry,
    paper      : ExamPaper,
    slot       : int,
    assignment : dict[str, int],
    paper_map  : dict[str, ExamPaper],
    exam_slots : list[tuple],
) -> bool:
    """
    Check whether placing *paper* at *slot* is compatible with difficulty
    constraints for papers already assigned on the same calendar day.

    Rules (same grade only):
      - Red   blocks Red + Yellow on the same day
      - Yellow blocks Red          on the same day
      - Green never blocks
    Linked partner papers are exempt from each other.
    """
    my_diff = registry.get_difficulty(paper.subject, paper.grade)
    if my_diff == "green":
        return True

    target_date = exam_slots[min(slot, len(exam_slots) - 1)][0]

    for lbl, s in assignment.items():
        other = paper_map[lbl]
        if other.grade != paper.grade:
            continue
        other_date = exam_slots[min(s, len(exam_slots) - 1)][0]
        if other_date != target_date:
            continue
        # Linked partners are exempt
        if lbl in paper.links or paper.label in other.links:
            continue
        other_diff = registry.get_difficulty(other.subject, other.grade)
        if other_diff == "green":
            continue
        if my_diff == "red" and other_diff in ("red", "yellow"):
            return False
        if my_diff == "yellow" and other_diff == "red":
            return False

    return True


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
    config           : CostConfig (duck-typed) — penalty weights; None uses defaults

    Returns
    -------
    ScheduleResult
    """
    papers = registry.all_papers()
    graph  = build_paper_clash_graph(papers)

    # ── Cost weights from config (duck-typed) or hardcoded defaults ──────────
    _w_day       = config.day_density_factor   if config else 5
    _w_week_base = config.week_density_base    if config else 6
    _w_same_week = config.same_week_penalty    if config else 1
    _w_teacher   = config.teacher_load_penalty if config else 1

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
    paper_map       : dict[str, ExamPaper] = {p.label: p for p in papers}

    # ── Step 0: Pinned papers (including ST) ─────────────────────────────────

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

    # Build grade-reserved slots from ST papers — these slots are blocked
    # for all non-ST papers in that grade
    grade_reserved_slots: dict[str, set[int]] = defaultdict(set)
    for p in pinned_papers:
        if p.subject == "ST":
            grade_reserved_slots[p.grade].add(assignment[p.label])

    # ── Step 0.5: Linked paper pre-placement ────────────────────────────────────
    #    Place linked pairs before priority / DSatur passes.
    #    Primary is placed in an AM slot (even index); partner is placed in the
    #    PM slot of the same day (primary_slot + 1).  Falls back if PM blocked.

    warnings: dict[str, list[str]] = {p.label: [] for p in papers}
    linked_label_set: set[str] = set()

    seen_pairs: set[frozenset[str]] = set()
    linked_pairs: list[tuple[ExamPaper, ExamPaper]] = []
    for p in papers:
        for link_label in p.links:
            pair_key = frozenset({p.label, link_label})
            if pair_key in seen_pairs:
                continue
            seen_pairs.add(pair_key)
            partner = paper_map.get(link_label)
            if partner is None:
                continue
            # Skip if either is already placed (pinned)
            if p.label in assignment or partner.label in assignment:
                continue
            if p.student_count() >= partner.student_count():
                linked_pairs.append((p, partner))
            else:
                linked_pairs.append((partner, p))

    linked_pairs.sort(
        key=lambda pair: pair[0].student_count() + pair[1].student_count(),
        reverse=True,
    )

    for primary, partner in linked_pairs:
        # Place primary at best valid AM slot (even index = AM)
        reserved = grade_reserved_slots.get(primary.grade, set())
        forbidden = {assignment[nb] for nb in graph[primary.label]
                     if nb in assignment}
        am_slots = [s for s in range(0, num_slots, 2)]   # even indices = AM
        valid = [s for s in am_slots
                 if s not in forbidden and s not in reserved
                 and _difficulty_allows(registry, primary, s, assignment,
                                        paper_map, exam_slots)]
        if not valid:
            # Relax difficulty constraint, still prefer AM
            valid = [s for s in am_slots
                     if s not in forbidden and s not in reserved] \
                    or am_slots or list(range(num_slots))
        best = _pick_spread_slot(valid, assignment, primary.label, num_slots)
        assignment[primary.label] = best
        linked_label_set.add(primary.label)

        # Place partner in the PM slot of the same day (best + 1)
        pm_slot = best + 1  # AM slot is even, so AM+1 is PM of same day
        reserved_p = grade_reserved_slots.get(partner.grade, set())
        forbidden_p = {assignment[nb] for nb in graph[partner.label]
                       if nb in assignment}
        if (pm_slot < num_slots
                and pm_slot not in forbidden_p
                and pm_slot not in reserved_p
                and _difficulty_allows(registry, partner, pm_slot,
                                       assignment, paper_map, exam_slots)):
            assignment[partner.label] = pm_slot
            linked_label_set.add(partner.label)
        else:
            # Fallback: place independently with a warning
            valid_p = [s for s in range(num_slots)
                       if s not in forbidden_p and s not in reserved_p
                       and _difficulty_allows(registry, partner, s, assignment,
                                              paper_map, exam_slots)]
            if not valid_p:
                valid_p = [s for s in range(num_slots)
                           if s not in forbidden_p and s not in reserved_p] \
                          or list(range(num_slots))
            fallback = _pick_spread_slot(valid_p, assignment, partner.label,
                                         num_slots)
            assignment[partner.label] = fallback
            linked_label_set.add(partner.label)
            warnings[partner.label].append(
                f"Linked to {primary.label} (slot {best + 1}) "
                f"— could not place in PM of same day, placed independently"
            )

    # ── Step 1: Priority placement ─────────────────────────────────────────────

    priority_papers = [
        p for p in unpinned_papers
        if p.subject in PRIORITY_SUBJECTS and p.label not in linked_label_set
    ]
    priority_papers.sort(
        key=lambda p: (-_grade_sort_key(p.grade), p.paper_number)
    )

    for paper in priority_papers:
        reserved = grade_reserved_slots.get(paper.grade, set())
        forbidden = {assignment[nb] for nb in graph[paper.label]
                     if nb in assignment}
        valid_slots = [s for s in range(num_slots)
                       if s not in forbidden and s not in reserved
                       and _difficulty_allows(registry, paper, s, assignment,
                                              paper_map, exam_slots)]
        if not valid_slots:
            valid_slots = [s for s in range(num_slots)
                           if s not in forbidden and s not in reserved] \
                          or list(range(num_slots))
        best = _pick_spread_slot(valid_slots, assignment, paper.label, num_slots)
        assignment[paper.label] = best
        priority_labels.add(paper.label)

    # ── Step 2: DSatur for remaining papers — spread across all slots ──────────

    remaining = [p for p in unpinned_papers
                 if p.label not in assignment]
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
        chosen_paper = unassigned[chosen_label]
        reserved = grade_reserved_slots.get(chosen_paper.grade, set())
        forbidden = {assignment[nb] for nb in graph[chosen_label]
                     if nb in assignment}
        valid = [s for s in range(num_slots)
                 if s not in forbidden and s not in reserved
                 and _difficulty_allows(registry, chosen_paper, s, assignment,
                                        paper_map, exam_slots)]
        if not valid:
            valid = [s for s in range(num_slots)
                     if s not in forbidden and s not in reserved] \
                    or list(range(num_slots))
        slot = _pick_spread_slot(valid, assignment, chosen_label, num_slots)
        assignment[chosen_label] = slot

        for nb in graph[chosen_label]:
            if nb in unassigned and nb != chosen_label:
                saturation[nb].add(slot)

        del unassigned[chosen_label]

    # ── Step 3: Spacing penalty pass ───────────────────────────────────────────

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
                    if _w_same_week == 0:
                        continue
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

    teacher_warnings: list[str] = []
    if exam_tree is not None:
        teacher_warnings = _teacher_marking_warnings(assignment, paper_map, exam_tree)

    # ── Step 5: Student-load hill-climb ────────────────────────────────────────

    def _dp(k: int) -> int:
        return k * (k - 1) * _w_day if k > 1 else 0

    def _wp(k: int) -> int:
        return ((k - 1) * (k + _w_week_base) // 2) * _w_same_week if k > 1 else 0

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

            hc_paper = paper_map[lbl]
            reserved = grade_reserved_slots.get(hc_paper.grade, set())
            # Temporarily remove self so _difficulty_allows sees other papers only
            del assignment[lbl]

            for sl_new in range(num_slots):
                if sl_new == sl_old or sl_new in forbidden:
                    continue
                if sl_new in reserved:
                    continue
                # Difficulty feasibility
                if not _difficulty_allows(registry, hc_paper, sl_new,
                                          assignment, paper_map, exam_slots):
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

                # AM-first tiebreak: equal-cost improvements prefer AM (even index)
                if delta < best_delta or (delta == best_delta and delta < 0
                                          and sl_new % 2 == 0):
                    best_delta = delta
                    best_slot  = sl_new

            # Restore self in assignment
            assignment[lbl] = sl_old

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

    # ── Build penalty log ────────────────────────────────────────────────────

    penalty_log: list[PenaltyEntry] = []

    # Day density penalties — group by (day, paper-set) to keep log manageable
    day_paper_groups: dict[date, list[str]] = defaultdict(list)
    for lbl, sl in assignment.items():
        day_paper_groups[slot_date(sl)].append(lbl)

    for sid, dc in s_day_c.items():
        for d, k in dc.items():
            if k > 1:
                day_papers = [
                    lbl for lbl in day_paper_groups.get(d, [])
                    if sid in paper_map[lbl].student_ids
                ]
                penalty_log.append(PenaltyEntry(
                    constraint="day_density",
                    papers=day_papers,
                    entity=str(sid),
                    value=_dp(k),
                ))

    # Week density penalties
    week_paper_groups: dict[int, list[str]] = defaultdict(list)
    for lbl, sl in assignment.items():
        week_paper_groups[_week_number(slot_date(sl))].append(lbl)

    for sid, wc in s_wk_c.items():
        for wk, k in wc.items():
            if k > 1:
                week_papers = [
                    lbl for lbl in week_paper_groups.get(wk, [])
                    if sid in paper_map[lbl].student_ids
                ]
                penalty_log.append(PenaltyEntry(
                    constraint="week_density",
                    papers=week_papers,
                    entity=str(sid),
                    value=_wp(k),
                ))

    # Teacher load penalties
    for msg in teacher_warnings:
        penalty_log.append(PenaltyEntry(
            constraint="teacher_load",
            papers=[],
            entity=msg,
            value=_w_teacher,
        ))

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
        penalty_log        = penalty_log,
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
