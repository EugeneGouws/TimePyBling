"""
exam_clash.py

Purpose
-------
Traverse the ExamTree and, per grade, find the MINIMUM number of exam
slots so that no student writes two subjects at the same time.

Algorithm: Two-phase exact colouring
-------------------------------------
Phase 1 — DSatur (upper bound)
    A greedy MRV algorithm that gives a good starting solution fast.
    For most school timetables this is already optimal.

Phase 2 — Backtracking (exact minimum)
    We try to beat the DSatur result by attempting to colour the graph
    with one fewer slot. If it succeeds we try again with one fewer.
    We stop when backtracking proves the current count is impossible.
    The last successful colouring is the true minimum.

Pruning techniques used in backtracking:
    - Forward checking: after assigning a slot to a node, immediately
      check whether any unassigned neighbour now has no valid slots left.
      If so, abandon this branch immediately.
    - Symmetry breaking: on the first node, we always assign slot 0.
      This prevents exploring equivalent permutations of slot labels.
    - Node ordering: process nodes in descending degree order so the
      most-constrained subjects are placed first.

For school-sized graphs (typically 15-35 subjects per grade) this runs
in well under a second.
"""

from reader.exam_tree import ExamTree


# ------------------------------------------------
# BUILD CLASH GRAPH
# ------------------------------------------------
def build_clash_graph(student_sets: dict) -> dict:
    """
    Build an adjacency set for each subject.

    Two subjects clash if their student sets have any overlap.

    Returns
    -------
    dict  { subject_label -> set of subject_labels it clashes with }
    """
    subjects = list(student_sets.keys())
    graph    = {s: set() for s in subjects}

    for i in range(len(subjects)):
        for j in range(i + 1, len(subjects)):
            a, b = subjects[i], subjects[j]
            if not student_sets[a].isdisjoint(student_sets[b]):
                graph[a].add(b)
                graph[b].add(a)

    return graph

def is_excluded(subject_label: str, exclusions: set) -> bool:
    """
    Returns True if this subject should be skipped.

    exclusions can contain:
        'LIB'    -> skips LIB for ALL grades
        'MU_08'  -> skips MU only in Grade 8

    Examples:
        is_excluded('LIB_08', {'LIB'})    -> True
        is_excluded('MU_08',  {'MU_08'})  -> True
        is_excluded('MU_09',  {'MU_08'})  -> False
    """
    code = subject_label.split("_")[0]          # e.g. 'LIB' from 'LIB_08'
    return code in exclusions or subject_label in exclusions

# ------------------------------------------------
# PHASE 1: DSATUR (upper bound)
# ------------------------------------------------
def dsatur_colouring(graph: dict) -> dict:
    """
    Assign slots using DSatur (Degree of Saturation / MRV).

    At each step, pick the uncoloured node with the most distinct
    slot numbers already used by its neighbours (saturation).
    Ties broken by degree.

    Returns
    -------
    dict  { subject_label -> slot_number (0-indexed) }
    """
    assignment = {}
    saturation = {s: set() for s in graph}
    uncoloured = set(graph.keys())

    while uncoloured:
        chosen = max(
            uncoloured,
            key=lambda s: (len(saturation[s]), len(graph[s]))
        )

        neighbour_slots = {
            assignment[n] for n in graph[chosen] if n in assignment
        }

        slot = 0
        while slot in neighbour_slots:
            slot += 1

        assignment[chosen] = slot
        uncoloured.remove(chosen)

        for neighbour in graph[chosen]:
            if neighbour in uncoloured:
                saturation[neighbour].add(slot)

    return assignment


# ------------------------------------------------
# PHASE 2: BACKTRACKING (exact minimum)
# ------------------------------------------------
def _backtrack(nodes, index, graph, assignment, num_slots):
    """
    Recursive backtracking with forward checking.

    Tries to colour all nodes using at most `num_slots` colours.
    Returns True if a valid colouring was found (assignment is updated).
    Returns False if it is impossible with this many slots.

    Parameters
    ----------
    nodes      : list of node labels in processing order
    index      : current position in nodes list
    graph      : adjacency sets
    assignment : dict being built (mutated in place)
    num_slots  : maximum number of slots allowed
    """
    if index == len(nodes):
        return True  # All nodes successfully coloured

    node = nodes[index]

    # Slots already used by neighbours that are already assigned
    forbidden = {
        assignment[n] for n in graph[node] if n in assignment
    }

    for slot in range(num_slots):
        if slot in forbidden:
            continue

        assignment[node] = slot

        # Forward checking: for each unassigned neighbour, check that
        # it still has at least one valid slot remaining
        pruned = False
        for neighbour in graph[node]:
            if neighbour in assignment:
                continue
            neighbour_forbidden = {
                assignment[n]
                for n in graph[neighbour]
                if n in assignment
            }
            if len(neighbour_forbidden) >= num_slots:
                # This neighbour has no valid slots left — dead end
                pruned = True
                break

        if not pruned:
            if _backtrack(nodes, index + 1, graph, assignment, num_slots):
                return True

        del assignment[node]

    return False  # No valid slot found for this node


def exact_colouring(graph: dict, upper_bound: int) -> dict:
    """
    Find the true minimum colouring by trying to beat the upper bound.

    Starting from upper_bound - 1, we try progressively fewer slots.
    Each attempt uses backtracking with forward checking.
    We stop when backtracking fails — the previous success is optimal.

    Nodes are ordered by descending degree so the most-constrained
    subjects are placed first (MRV ordering).

    Parameters
    ----------
    graph       : adjacency sets from build_clash_graph()
    upper_bound : number of slots from DSatur (known to be achievable)

    Returns
    -------
    dict  { subject_label -> slot_number (0-indexed) }
          using the minimum number of slots found
    """
    # Order nodes by descending degree for better pruning
    nodes = sorted(graph.keys(), key=lambda n: len(graph[n]), reverse=True)

    # Start with the DSatur result — guaranteed to be valid
    best_assignment = dsatur_colouring(graph)
    best_slots      = upper_bound

    # Try to beat it one slot at a time
    for target in range(upper_bound - 1, 0, -1):
        attempt = {}

        # Symmetry breaking: fix the first node to slot 0
        if nodes:
            attempt[nodes[0]] = 0
            start_index = 1
        else:
            start_index = 0

        if _backtrack(nodes, start_index, graph, attempt, target):
            best_assignment = attempt
            best_slots      = target
        else:
            # Proven impossible — best_slots is the true minimum
            break

    return best_assignment


# ------------------------------------------------
# CLASH REPORT
# ------------------------------------------------
def print_clash_report(exam_tree: ExamTree, exclusions: set = None):
    """
    Print the full exam slot report for every grade.

    Parameters
    ----------
    exam_tree  : built ExamTree
    exclusions : set of subject codes to skip (e.g. {'ST', 'LIB', 'PE'})
    """
    if exclusions is None:
        exclusions = set()

    print("=" * 60)
    print("EXAM SLOT REPORT  (exact minimum via backtracking)")
    print("Subjects in the same slot share NO students")
    print("=" * 60)
    print()

    for grade_label in sorted(exam_tree.grades.keys()):
        grade_node = exam_tree.grades[grade_label]

        student_sets = {
            label: subject.all_students()
            for label, subject in grade_node.exam_subjects.items()
            if not is_excluded(label, exclusions)
        }

        if not student_sets:
            continue

        graph       = build_clash_graph(student_sets)
        dsatur_res  = dsatur_colouring(graph)
        upper_bound = max(dsatur_res.values()) + 1
        assignment  = exact_colouring(graph, upper_bound)
        num_slots   = max(assignment.values()) + 1

        # Group subjects by slot
        slots = {}
        for subj, slot in assignment.items():
            slots.setdefault(slot, []).append(subj)

        # Sort subjects within each slot
        for slot in slots.values():
            slot.sort()

        # Sort slots by number of students writing (busiest first)
        sorted_slots = sorted(
            slots.values(),
            key=lambda group: sum(len(student_sets[s]) for s in group),
            reverse=True
        )

        print(f"{grade_label}  ({num_slots} slots needed)")

        for i, group in enumerate(sorted_slots, start=1):
            slot_students = set()
            for s in group:
                slot_students |= student_sets[s]
            print(
                f"  Slot {i:>2}  |  {len(group):>2} subject(s)"
                f"  |  {len(slot_students):>3} students writing"
                f"  |  {', '.join(group)}"
            )

        print()

    print("-" * 60)
    print("SUMMARY")
    print("-" * 60)

    for grade_label in sorted(exam_tree.grades.keys()):
        grade_node = exam_tree.grades[grade_label]

        student_sets = {
            label: subject.all_students()
            for label, subject in grade_node.exam_subjects.items()
            if not is_excluded(label, exclusions)
        }

        if not student_sets:
            continue

        graph       = build_clash_graph(student_sets)
        dsatur_res  = dsatur_colouring(graph)
        upper_bound = max(dsatur_res.values()) + 1
        assignment  = exact_colouring(graph, upper_bound)
        num_slots   = max(assignment.values()) + 1

        print(f"  {grade_label}: {num_slots} slots")

    print()


# ------------------------------------------------
# OPTIONAL TEST NOTE
# ------------------------------------------------
if __name__ == "__main__":
    print("This file is meant to be used from main.py")