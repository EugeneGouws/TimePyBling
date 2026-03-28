"""
cpsat_optimiser.py
------------------
Global exam-schedule optimiser using OR-Tools CP-SAT solver.

Minimises teacher marking load subject to a student-cost upper-bound
constraint and hard constraints (student clashes, AM-before-PM,
pinned papers, grade-reserved slots).

Not currently called from the main flow — the controller uses
hill_climb_teacher instead.  This module is retained for future use
when exact optimality proof is needed.
"""

from __future__ import annotations

from collections import defaultdict

try:
    from ortools.sat.python import cp_model
    HAS_ORTOOLS = True
except ImportError:
    HAS_ORTOOLS = False

from core.cost_function import (
    StudentStressCost, TeacherMarkingCost,
    _colour_weight, _max_window_coeff, _PASSES, _SESSIONS,
)
from reader.exam_paper import ExamPaperRegistry
from reader.exam_clash import build_paper_clash_graph


def cpsat_optimise_teacher(
    schedule_dict: dict,
    student_cost_fn: StudentStressCost,
    teacher_cost_fn: TeacherMarkingCost,
    registry: ExamPaperRegistry,
    exam_tree,
    student_baseline: float,
    teacher_tolerance_pct: int,
    time_limit_seconds: int = 30,
) -> tuple[dict, float, float, bool]:
    """Minimise teacher cost subject to student cost upper bound.

    Returns (optimised_schedule, final_student_cost, final_teacher_cost, is_optimal).
    """
    if not HAS_ORTOOLS:
        raise ImportError("ortools is not installed")

    # ── Collect papers and slots ─────────────────────────────────────────
    all_papers = []
    paper_by_label: dict[str, object] = {}
    for papers in schedule_dict.values():
        for p in papers:
            if p.label not in paper_by_label:
                paper_by_label[p.label] = p
                all_papers.append(p)

    if not all_papers:
        return schedule_dict, 0.0, 0.0, True

    all_keys = sorted(schedule_dict.keys())
    num_slots = len(all_keys)
    slot_to_idx = {k: i for i, k in enumerate(all_keys)}
    idx_to_key = {i: k for k, i in slot_to_idx.items()}
    max_day = max(d for d, _ in all_keys)
    num_days = max_day + 1

    slot_day = {i: k[0] for i, k in enumerate(all_keys)}

    # ── Build grade-reserved slots ─────────────────────────────────────
    grade_reserved: dict[str, set[int]] = defaultdict(set)
    for p in all_papers:
        if p.subject == "ST" and p.pinned_slot is not None:
            grade_reserved[p.grade].add(p.pinned_slot)

    # ── Build clash graph ──────────────────────────────────────────────
    graph = build_paper_clash_graph(all_papers)

    # ── Build model ────────────────────────────────────────────────────
    model = cp_model.CpModel()

    x: dict[str, dict[int, object]] = {}
    for p in all_papers:
        x[p.label] = {}
        for s in range(num_slots):
            x[p.label][s] = model.new_bool_var(f"x_{p.label}_{s}")

    # C1: Each paper in exactly one slot
    for p in all_papers:
        model.add_exactly_one(x[p.label][s] for s in range(num_slots))

    # C2: Clashing papers not in same slot
    seen_pairs: set[frozenset] = set()
    for label_a, neighbours in graph.items():
        if label_a not in x:
            continue
        for label_b in neighbours:
            if label_b not in x:
                continue
            pair = frozenset({label_a, label_b})
            if pair in seen_pairs:
                continue
            seen_pairs.add(pair)
            for s in range(num_slots):
                model.add(x[label_a][s] + x[label_b][s] <= 1)

    # C3: AM-before-PM
    for d in range(num_days):
        am_key = (d, "AM")
        pm_key = (d, "PM")
        if am_key not in slot_to_idx or pm_key not in slot_to_idx:
            continue
        am_idx = slot_to_idx[am_key]
        pm_idx = slot_to_idx[pm_key]
        am_sum = sum(x[p.label][am_idx] for p in all_papers)
        pm_sum = sum(x[p.label][pm_idx] for p in all_papers)
        model.add(pm_sum <= len(all_papers) * am_sum)

    # C4: Pinned papers fixed
    for p in all_papers:
        if p.pinned_slot is not None:
            pinned_clamped = min(p.pinned_slot, num_slots - 1)
            pinned_day = pinned_clamped // 2
            pinned_session = "AM" if pinned_clamped % 2 == 0 else "PM"
            pinned_key = (pinned_day, pinned_session)
            if pinned_key in slot_to_idx:
                model.add(x[p.label][slot_to_idx[pinned_key]] == 1)

    # C5: Grade-reserved slots
    for p in all_papers:
        if p.subject == "ST":
            continue
        reserved = grade_reserved.get(p.grade, set())
        for raw_slot in reserved:
            rd = raw_slot // 2
            rs = "AM" if raw_slot % 2 == 0 else "PM"
            rk = (rd, rs)
            if rk in slot_to_idx:
                model.add(x[p.label][slot_to_idx[rk]] == 0)

    # ── Objective: minimise teacher marking cost (with position ramp) ──
    SCALE = 10_000
    max_wc = _max_window_coeff(max_day)

    teacher_map = teacher_cost_fn._teacher_map
    marking_total_base = teacher_cost_fn._total_base
    marking_max_raw = marking_total_base * max_wc if marking_total_base else 1

    # Per-slot marking coefficient with position ramp
    marking_slot_coeff: dict[int, float] = {}
    for s in range(num_slots):
        d = slot_day[s]
        total_w = 0.0
        for w_size, w_weight in _PASSES:
            upper = max_day - w_size + 1
            if upper < 0:
                continue
            lo = max(0, d - w_size + 1)
            hi = min(d, upper)
            for d_start in range(lo, hi + 1):
                position_weight = (d_start + 1) / num_days
                total_w += w_weight * position_weight
        marking_slot_coeff[s] = total_w

    marking_terms = []
    for p in all_papers:
        teachers = teacher_map.get(p.label, {})
        if not teachers:
            continue
        total_scripts = sum(teachers.values())
        for s in range(num_slots):
            coeff = int(total_scripts * marking_slot_coeff[s] * SCALE / marking_max_raw)
            if coeff > 0:
                marking_terms.append(coeff * x[p.label][s])

    if marking_terms:
        model.minimize(sum(marking_terms))

    # ── Solve ──────────────────────────────────────────────────────────
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_seconds
    solver.parameters.num_workers = 4

    status = solver.solve(model)

    is_optimal = (status == cp_model.OPTIMAL)
    is_feasible = status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    if not is_feasible:
        s_cost = student_cost_fn.compute(schedule_dict)
        t_cost = teacher_cost_fn.compute(schedule_dict)
        return schedule_dict, s_cost, t_cost, False

    # ── Extract solution ───────────────────────────────────────────────
    new_schedule: dict = {k: [] for k in all_keys}
    for p in all_papers:
        for s in range(num_slots):
            if solver.value(x[p.label][s]) == 1:
                new_schedule[idx_to_key[s]].append(p)
                break

    final_student = student_cost_fn.compute(new_schedule)
    final_teacher = teacher_cost_fn.compute(new_schedule)

    # Reject if student cost exceeds tolerance
    max_student = student_baseline * (1 + teacher_tolerance_pct / 100)
    if final_student > max_student:
        s_cost = student_cost_fn.compute(schedule_dict)
        t_cost = teacher_cost_fn.compute(schedule_dict)
        return schedule_dict, s_cost, t_cost, False

    return new_schedule, final_student, final_teacher, is_optimal
