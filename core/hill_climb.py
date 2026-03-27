"""
hill_climb.py
-------------
Slot-swap hill-climb optimiser over an exam schedule.

A "move" is a slot-swap: two (day, session) slots exchange their paper lists,
including swapping with an empty slot.  Pinned papers are never moved.

Delta cost is computed by recomputing only the windows that contain at least
one of the two swapped days, not the full schedule.

Public API
----------
hill_climb(schedule, cost_fn, max_iterations) -> (schedule, final_cost)
"""

from __future__ import annotations

from core.cost_function import TotalCost

_PASSES = [(2, 3), (3, 2), (4, 1)]


def _affected_windows(
    d1: int, d2: int, max_day: int
) -> list[tuple[int, int, int]]:
    """
    Return (window_start, window_size, window_weight) for every window that
    contains d1 or d2, deduplicated.

    A window of size W starting at day d contains day x iff d <= x <= d+W-1,
    i.e. x-W+1 <= d <= x, clamped to [0, max_day-W+1].
    """
    seen:   set[tuple[int, int]]          = set()
    result: list[tuple[int, int, int]]    = []
    for w_size, w_weight in _PASSES:
        upper = max_day - w_size + 1
        if upper < 0:
            continue
        for day in (d1, d2):
            lo = max(0, day - w_size + 1)
            hi = min(day, upper)
            for d in range(lo, hi + 1):
                key = (d, w_size)
                if key not in seen:
                    seen.add(key)
                    result.append((d, w_size, w_weight))
    return result


def _has_pinned(papers: list) -> bool:
    """True if any paper in the list is pinned."""
    return any(p.pinned_slot is not None for p in papers)


def hill_climb(
    schedule: dict,
    cost_fn: TotalCost,
    max_iterations: int = 10_000,
) -> tuple[dict, float]:
    """
    Hill-climb over the schedule by swapping pairs of slots.

    Parameters
    ----------
    schedule : dict
        {(day: int, session: str): list[ExamPaper]} — not mutated.
    cost_fn : TotalCost
    max_iterations : int
        Maximum number of full improving passes (not individual swaps).

    Returns
    -------
    (improved_schedule, final_cost)
        improved_schedule has the same key set as the input.
    """
    # Shallow-copy: lists are copied so the caller's dict is not mutated
    sched = {k: list(v) for k, v in schedule.items()}

    if not sched:
        return sched, 0.0

    max_day  = max(d for d, _ in sched)
    all_keys = sorted(sched.keys())

    for _iteration in range(max_iterations):
        improved = False

        for i, k1 in enumerate(all_keys):
            d1, _ = k1

            for k2 in all_keys[i + 1:]:
                d2, _ = k2

                papers1 = sched[k1]
                papers2 = sched[k2]

                # Never move pinned papers
                if _has_pinned(papers1) or _has_pinned(papers2):
                    continue

                # Nothing to do if both slots are empty
                if not papers1 and not papers2:
                    continue

                affected    = _affected_windows(d1, d2, max_day)
                cost_before = cost_fn.compute_windows(sched, affected)

                sched[k1], sched[k2] = sched[k2], sched[k1]

                cost_after = cost_fn.compute_windows(sched, affected)

                if cost_after - cost_before < 0:
                    improved = True
                else:
                    # Revert
                    sched[k1], sched[k2] = sched[k2], sched[k1]

        if not improved:
            break

    return sched, cost_fn.compute(sched)
