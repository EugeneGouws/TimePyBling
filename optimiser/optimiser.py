"""
optimiser.py
------------
Simulated annealing optimiser for the school timetable.

Public API
----------
    SAConfig   — dataclass of SA hyperparameters
    SAResult   — result object passed to done_cb
    SARunner   — runs the SA on a background thread; called by the UI

Move set
--------
    swap_blocks(a, b)  — atomically swaps all 7 subblock contents between
                         two block letters.  This is the ONLY legal SA move.
                         Calling it twice reverts the state, so reject is free.

Seeding
-------
    The SA receives an already-converted BlockTree (self.block_tree from the UI).
    It works on an internal deep-copy; at the end it writes the best solution
    back into the original bt.blocks in-place so the UI's cost panel refreshes
    correctly without needing to reload the file.
"""

import copy
import math
import random
import threading
from dataclasses import dataclass

from core.block_tree          import BlockTree
from optimiser.cost_function  import evaluate, CostConfig


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class SAConfig:
    """Hyperparameters for the simulated annealing run."""
    T_start      : float = 1_000.0
    T_min        : float = 0.1
    cooling_rate : float = 0.9999
    max_iter     : int   = 500_000
    log_every    : int   = 5_000     # how often to emit a progress line


# ─────────────────────────────────────────────────────────────────────────────
# RESULT
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class SAResult:
    """
    Returned (via done_cb) when the SA finishes or is stopped.

    Attributes
    ----------
    initial_cost  : E(T) of the timetable before SA started
    best_cost     : lowest E(T) seen during the run
    final_cost    : E(T) at the moment the run ended (may be > best_cost
                    if the run was stopped mid-improvement)
    iterations    : number of moves attempted
    accepted      : number of accepted moves
    improved      : True if best_cost < initial_cost
    stopped_early : True if stop() was called before the run completed
    """
    initial_cost  : float
    best_cost     : float
    final_cost    : float
    iterations    : int
    accepted      : int
    improved      : bool
    stopped_early : bool

    def summary(self) -> str:
        direction = "↓" if self.improved else "→"
        pct = (
            100.0 * (self.initial_cost - self.best_cost) / self.initial_cost
            if self.initial_cost > 0 else 0.0
        )
        lines = [
            f"{'─' * 44}",
            f"  SA result",
            f"{'─' * 44}",
            f"  Initial cost   : {self.initial_cost:>10.1f}",
            f"  Best cost      : {self.best_cost:>10.1f}  {direction}  ({pct:.2f}% improvement)",
            f"  Iterations     : {self.iterations:>10,}",
            f"  Accepted moves : {self.accepted:>10,}",
            f"{'─' * 44}",
        ]
        if self.stopped_early:
            lines.append("  (run stopped early by user)")
        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# RUNNER
# ─────────────────────────────────────────────────────────────────────────────

class SARunner:
    """
    Runs the SA optimiser on a background thread.

    Parameters
    ----------
    bt          : the BlockTree already loaded in the UI (self.block_tree).
                  The SA works on an internal deep-copy; at the end the best
                  solution is written back into bt.blocks in-place so the
                  UI's cost panel reflects the improved timetable automatically.
    config      : SAConfig — all defaults are sensible
    progress_cb : callable(msg: str) — called periodically with a log line.
                  Invoked from the background thread; the UI wraps it with
                  self.after(0, ...) so tkinter safety is handled there.
    done_cb     : callable(SAResult) — called once when the run finishes or
                  is stopped.  Also invoked from the background thread.
    cost_config : CostConfig passed to evaluate(); None uses defaults.
    """

    def __init__(self,
                 bt          : BlockTree,
                 config      : SAConfig   = None,
                 progress_cb             = None,
                 done_cb                 = None,
                 cost_config : CostConfig = None):

        self._bt          = bt
        self._config      = config      or SAConfig()
        self._progress_cb = progress_cb or (lambda msg: None)
        self._done_cb     = done_cb     or (lambda result: None)
        self._cost_config = cost_config or CostConfig()

        self._stop_event = threading.Event()
        self._thread     = None

    # ------------------------------------------------------------------
    # Public control API
    # ------------------------------------------------------------------

    def start(self):
        """Launch the SA on a daemon background thread."""
        if self.is_running():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        """Signal the SA to stop after the current iteration."""
        self._stop_event.set()

    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    # ------------------------------------------------------------------
    # SA loop  (runs on background thread)
    # ------------------------------------------------------------------

    def _run(self):
        cfg  = self._config
        log  = self._progress_cb
        stop = self._stop_event

        # Work on a deep copy — never mutate the UI's bt mid-run
        bt          = copy.deepcopy(self._bt)
        block_names = list(bt.blocks.keys())

        initial_score = evaluate(bt, self._cost_config).total
        score         = initial_score
        best_bt       = copy.deepcopy(bt)
        best_score    = score

        T          = cfg.T_start
        iteration  = 0
        accepted   = 0

        log(f"SA started  |  blocks: {block_names}  |  initial E(T): {initial_score:.1f}")
        log(f"Config: T_start={cfg.T_start}  T_min={cfg.T_min}  "
            f"cooling={cfg.cooling_rate}  max_iter={cfg.max_iter:,}")
        log("─" * 60)

        while T > cfg.T_min and iteration < cfg.max_iter:

            if stop.is_set():
                break

            # ── Propose: two distinct block letters ──────────────────
            a, b = random.sample(block_names, 2)

            # ── Apply ────────────────────────────────────────────────
            bt.swap_blocks(a, b)
            new_score = evaluate(bt, self._cost_config).total
            delta     = new_score - score

            # ── Accept / reject ──────────────────────────────────────
            if delta < 0 or random.random() < math.exp(-delta / T):
                score    = new_score
                accepted += 1
                if score < best_score:
                    best_score = score
                    best_bt    = copy.deepcopy(bt)
            else:
                bt.swap_blocks(a, b)   # revert — swap is its own inverse

            T         *= cfg.cooling_rate
            iteration += 1

            # ── Periodic progress line ────────────────────────────────
            if iteration % cfg.log_every == 0:
                log(f"  iter {iteration:>7,}  |  T={T:>8.3f}  "
                    f"|  current={score:>8.1f}  |  best={best_score:>8.1f}")

        # ── Write best solution back into the original bt in-place ───
        # Updates self.block_tree in the UI without a separate assignment.
        # The Export tab and cost panel will both see the improved timetable.
        self._bt.blocks = best_bt.blocks

        stopped_early = stop.is_set() and iteration < cfg.max_iter
        result = SAResult(
            initial_cost  = initial_score,
            best_cost     = best_score,
            final_cost    = score,
            iterations    = iteration,
            accepted      = accepted,
            improved      = best_score < initial_score,
            stopped_early = stopped_early,
        )

        log("─" * 60)
        log(f"SA finished  |  {iteration:,} iterations  |  "
            f"{initial_score:.1f} → {best_score:.1f}")

        self._done_cb(result)