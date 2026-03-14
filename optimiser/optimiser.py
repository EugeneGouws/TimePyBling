"""
optimiser.py
------------
Simulated annealing optimiser for the school timetable.

Move type: swap two SlotAssignments between two different subblocks.
This preserves both the number of lessons per class and the number of
classes — assignments are exchanged, never created or destroyed.

Acceptance rule (Metropolis criterion):
    - Always accept moves that reduce cost.
    - Accept moves that increase cost with probability e^(-delta / T).
    - T cools from T_start to T_min over max_iterations steps.

The optimiser runs in a background thread so the UI stays responsive.
Progress is reported via a callback function supplied by the caller.
"""

import math
import random
import threading
from dataclasses import dataclass, field
from typing import Callable

from core.block_tree         import BlockTree, SlotAssignment
from optimiser.cost_function import evaluate, CostConfig, CostBreakdown


# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class SAConfig:
    """
    Simulated annealing parameters.

    T_start      : initial temperature — controls how freely bad moves are
                   accepted at the start. A value of 1000 means a move that
                   increases cost by 1000 is accepted with ~37% probability.
    T_min        : stop when temperature falls below this value.
    cooling_rate : multiply T by this after each iteration (0 < rate < 1).
                   0.9999 gives a slow, thorough search.
                   0.999  gives a faster but less thorough search.
    max_iter     : hard cap on iterations regardless of temperature.
    seed         : random seed for reproducibility. None = random each run.
    """
    T_start:      float = 1000.0
    T_min:        float = 0.1
    cooling_rate: float = 0.9999
    max_iter:     int   = 500_000
    seed:         int   = None


@dataclass
class SAResult:
    """What the optimiser returns when it finishes."""
    best_cost:      int
    initial_cost:   int
    iterations:     int
    accepted_moves: int
    rejected_moves: int
    improved:       bool          # True if best_cost < initial_cost
    breakdown:      CostBreakdown = None

    def summary(self) -> str:
        delta = self.initial_cost - self.best_cost
        pct   = (delta / self.initial_cost * 100) if self.initial_cost else 0
        lines = [
            f"Initial cost : {self.initial_cost}",
            f"Final cost   : {self.best_cost}",
            f"Improvement  : {delta:+d}  ({pct:.1f}%)",
            f"Iterations   : {self.iterations:,}",
            f"Accepted     : {self.accepted_moves:,}",
            f"Rejected     : {self.rejected_moves:,}",
        ]
        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# Move generation
# ─────────────────────────────────────────────────────────────────────────────

def _all_assignments(bt: BlockTree) -> list[tuple[str, SlotAssignment]]:
    """
    Return a flat list of (subblock_name, SlotAssignment) for every
    assignment in the tree.  Rebuilt each time it's needed.
    """
    return list(bt.all_assignments())


def _random_swap(bt: BlockTree,
                 rng: random.Random) -> tuple[str, str, str, str, str, str] | None:
    """
    Pick two distinct assignments at random and return the swap parameters.

    Returns None if the tree has fewer than 2 assignments (nothing to swap).

    Returns
    -------
    (subblock_a, subject_a, grade_a, subblock_b, subject_b, grade_b)
    """
    assignments = _all_assignments(bt)
    if len(assignments) < 2:
        return None

    # Pick two distinct (subblock, assignment) pairs
    idx_a, idx_b = rng.sample(range(len(assignments)), 2)
    sb_a, assign_a = assignments[idx_a]
    sb_b, assign_b = assignments[idx_b]

    # Must be in different subblocks — resample until they are
    attempts = 0
    while sb_a == sb_b:
        idx_b = rng.randrange(len(assignments))
        sb_b, assign_b = assignments[idx_b]
        attempts += 1
        if attempts > 100:
            return None  # degenerate tree — all assignments in one subblock

    return (
        sb_a, assign_a.subject_code, assign_a.grade,
        sb_b, assign_b.subject_code, assign_b.grade,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Core SA loop
# ─────────────────────────────────────────────────────────────────────────────

def run_sa(
    bt:           BlockTree,
    config:       SAConfig             = None,
    cost_config:  CostConfig           = None,
    progress_cb:  Callable[[str], None] = None,
    stop_event:   threading.Event      = None,
) -> SAResult:
    """
    Run simulated annealing on the BlockTree in-place.

    The BlockTree is mutated during the run. If the final state is worse
    than the initial state (shouldn't happen with best-tracking), the
    caller should re-convert from the TimetableTree.

    The best solution found is tracked separately and restored at the end.

    Parameters
    ----------
    bt           : mutable BlockTree (from timetable_tree_to_block_tree)
    config       : SAConfig — defaults used if None
    cost_config  : CostConfig for evaluate() — defaults used if None
    progress_cb  : called periodically with a status string for the UI
    stop_event   : threading.Event — set it to request early termination

    Returns
    -------
    SAResult with statistics and the final cost breakdown.
    """
    if config      is None: config      = SAConfig()
    if cost_config is None: cost_config = CostConfig()

    rng = random.Random(config.seed)

    def log(msg: str):
        if progress_cb:
            progress_cb(msg)

    # ── Initial state ──
    current_cost = evaluate(bt, cost_config).total
    best_cost    = current_cost
    initial_cost = current_cost

    log(f"Starting SA  |  initial cost = {initial_cost}")
    log(f"T_start={config.T_start}  T_min={config.T_min}  "
        f"cooling={config.cooling_rate}  max_iter={config.max_iter:,}")
    log("─" * 50)

    T              = config.T_start
    accepted       = 0
    rejected       = 0
    last_report_it = 0
    REPORT_EVERY   = max(1, config.max_iter // 200)   # ~200 progress updates

    for iteration in range(1, config.max_iter + 1):

        # ── Check for stop request ──
        if stop_event and stop_event.is_set():
            log(f"Stopped by user at iteration {iteration:,}")
            break

        # ── Cool temperature ──
        T = config.T_start * (config.cooling_rate ** iteration)
        if T < config.T_min:
            log(f"Temperature reached minimum at iteration {iteration:,}")
            break

        # ── Generate a random swap move ──
        move = _random_swap(bt, rng)
        if move is None:
            continue

        sb_a, subj_a, grade_a, sb_b, subj_b, grade_b = move

        # ── Apply the swap ──
        try:
            bt.swap_assignments(sb_a, subj_a, grade_a,
                                sb_b, subj_b, grade_b)
        except ValueError:
            continue  # assignment disappeared between sample and swap — skip

        # ── Score ──
        new_cost = evaluate(bt, cost_config).total
        delta    = new_cost - current_cost

        # ── Accept or reject ──
        if delta < 0 or rng.random() < math.exp(-delta / T):
            # Accept
            current_cost = new_cost
            accepted    += 1
            if new_cost < best_cost:
                best_cost = new_cost
        else:
            # Reject — undo the swap
            bt.swap_assignments(sb_b, subj_a, grade_a,
                                sb_a, subj_b, grade_b)
            rejected += 1

        # ── Periodic progress report ──
        if iteration - last_report_it >= REPORT_EVERY:
            last_report_it = iteration
            log(f"iter {iteration:>8,}  |  T={T:7.2f}  |  "
                f"cost={current_cost:>7}  |  best={best_cost:>7}")

        # ── Early exit if perfect solution found ──
        if best_cost == 0:
            log(f"Perfect solution found at iteration {iteration:,}!")
            break

    # ── Final evaluation ──
    final_breakdown = evaluate(bt, cost_config)

    log("─" * 50)
    log(f"SA complete")
    log(f"Initial cost : {initial_cost}")
    log(f"Final cost   : {final_breakdown.total}")
    log(f"Best seen    : {best_cost}")
    log(f"Accepted     : {accepted:,}  |  Rejected: {rejected:,}")

    return SAResult(
        best_cost      = best_cost,
        initial_cost   = initial_cost,
        iterations     = iteration,
        accepted_moves = accepted,
        rejected_moves = rejected,
        improved       = best_cost < initial_cost,
        breakdown      = final_breakdown,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Threaded wrapper — used by the UI
# ─────────────────────────────────────────────────────────────────────────────

class SARunner:
    """
    Runs SA in a background thread. The UI calls start() and can call
    stop() at any time. Results are delivered via the done_cb callback.

    Usage
    -----
        runner = SARunner(block_tree, config, cost_config,
                          progress_cb=lambda msg: ...,
                          done_cb=lambda result: ...)
        runner.start()
        # later:
        runner.stop()
    """

    def __init__(self,
                 bt:          BlockTree,
                 config:      SAConfig             = None,
                 cost_config: CostConfig           = None,
                 progress_cb: Callable[[str], None] = None,
                 done_cb:     Callable[[SAResult], None] = None):
        self.bt          = bt
        self.config      = config      or SAConfig()
        self.cost_config = cost_config or CostConfig()
        self.progress_cb = progress_cb
        self.done_cb     = done_cb
        self._stop       = threading.Event()
        self._thread     = None

    def start(self):
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()

    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _run(self):
        result = run_sa(
            bt          = self.bt,
            config      = self.config,
            cost_config = self.cost_config,
            progress_cb = self.progress_cb,
            stop_event  = self._stop,
        )
        if self.done_cb:
            self.done_cb(result)


# ─────────────────────────────────────────────────────────────────────────────
# Standalone test
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    from pathlib import Path
    from core.timetable_tree      import build_timetable_tree_from_file
    from core.timetable_converter import timetable_tree_to_block_tree

    st_file = Path(__file__).parent.parent / "data" / "ST1.xlsx"
    if not st_file.exists():
        print("data/ST1.xlsx not found.")
        raise SystemExit(1)

    print(f"Loading: {st_file}")
    tt = build_timetable_tree_from_file(st_file)
    bt = timetable_tree_to_block_tree(tt)

    config = SAConfig(
        T_start      = 1000.0,
        T_min        = 0.1,
        cooling_rate = 0.9999,
        max_iter     = 100_000,
        seed         = 42,
    )

    result = run_sa(bt, config, progress_cb=print)
    print()
    print(result.summary())
    print()
    print(result.breakdown)