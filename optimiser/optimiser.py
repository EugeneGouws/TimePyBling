"""
optimiser.py
------------
Simulated annealing optimiser for the school timetable.

Move set (two legal moves only):
    Major move : swap_blocks(a, b)        — swaps all 7 subblots of block A with B
    Minor move : swap_subblock_contents   — swaps two individual slots within a block
                 (only useful for mixed blocks like H; skip for now)

The SA seeds from the real timetable via timetable_tree_to_block_tree(),
so it always starts from a valid, fully-populated state.
"""

import random
import math
import copy
from pathlib import Path

from core.timetable_tree    import build_timetable_tree_from_file
from core.timetable_converter import timetable_tree_to_block_tree
from core.block_tree         import BlockTree
from optimiser.cost_function import evaluate
from optimiser.block_exporter import export_to_xlsx


def optimise(
    st1_path      : Path,
    output_path   : Path,
    T_start       : float = 1000.0,
    T_min         : float = 1.0,
    cooling_rate  : float = 0.995,
    max_iterations: int   = 10_000,
    progress_cb         = None,   # optional: callable(iteration, score, T)
) -> tuple[BlockTree, float]:
    """
    Run simulated annealing on the timetable loaded from st1_path.

    Returns (best_block_tree, best_score).
    Exports the best solution to output_path automatically.
    """

    # ------------------------------------------------------------------ #
    # 1. Seed from the real timetable — NEVER build from scratch
    # ------------------------------------------------------------------ #
    tt = build_timetable_tree_from_file(st1_path)
    bt = timetable_tree_to_block_tree(tt)

    score      = evaluate(bt)
    best_bt    = copy.deepcopy(bt)
    best_score = score

    block_names = list(bt.blocks.keys())   # ["A","B","C","D","E","F","G","H"]
    T           = T_start
    iteration   = 0

    # ------------------------------------------------------------------ #
    # 2. SA loop
    # ------------------------------------------------------------------ #
    while T > T_min and iteration < max_iterations:

        # Propose: two distinct random block letters
        a, b = random.sample(block_names, 2)

        # Apply move
        bt.swap_blocks(a, b)
        new_score = evaluate(bt)
        delta     = new_score - score

        # Accept or reject
        if delta < 0 or random.random() < math.exp(-delta / T):
            score = new_score                   # keep
            if score < best_score:
                best_score = score
                best_bt    = copy.deepcopy(bt)
        else:
            bt.swap_blocks(a, b)               # revert — swap is its own inverse

        T         *= cooling_rate
        iteration += 1

        if progress_cb:
            progress_cb(iteration, score, T)

    # ------------------------------------------------------------------ #
    # 3. Export the best solution found
    # ------------------------------------------------------------------ #
    export_to_xlsx(best_bt, str(output_path))
    print(f"Done. Best score: {best_score:.1f}  (started: {evaluate(timetable_tree_to_block_tree(tt)):.1f})")

    return best_bt, best_score