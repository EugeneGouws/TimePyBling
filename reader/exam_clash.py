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
    We stop when backtracking proves the current count is impossible,
    or when the node counter exceeds MAX_ATTEMPTS (timeout).
    The last successful colouring is the true minimum.

    Return values from _backtrack:
        True  — valid colouring found
        False — proven impossible with this many slots
        None  — timed out; result is neither confirmed nor ruled out

Pruning techniques used in backtracking:
    - Forward checking: after assigning a slot to a node, immediately
      check whether any unassigned neighbour now has no valid slots left.
    - Symmetry breaking: on the first node, always assign slot 0.
    - Node ordering: descending degree (most-constrained first).
"""

from reader.exam_tree import ExamTree

# Maximum backtrack nodes explored before falling back to DSatur result.
# 100 000 is ample for any grade that has an exact solution reachable quickly;
# grades where it binds will still return the DSatur result, which is optimal
# or within one slot of optimal in practice.
MAX_ATTEMPTS = 100_000


# ------------------------------------------------
# BUILD CLASH GRAPH
# ------------------------------------------------
def build_clash_graph(student_sets: dict) -> dict:
    """
    Build an adjacency set for each subject.
    Two subjects clash if their student sets have any overlap.
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
    """
    code = subject_label.split("_")[0]
    return code in exclusions or subject_label in exclusions


# ------------------------------------------------
# PHASE 1: DSATUR (upper bound)
# ------------------------------------------------
def dsatur_colouring(graph: dict) -> dict:
    """
    Assign slots using DSatur (Degree of Saturation / MRV).

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
def _backtrack(nodes, index, graph, assignment, num_slots, counter):
    """
    Recursive backtracking with forward checking and a node counter.

    Returns
    -------
    True   — valid colouring found
    False  — proven impossible with num_slots colours
    None   — node limit reached; result is inconclusive
    """
    counter[0] += 1
    if counter[0] > MAX_ATTEMPTS:
        return None  # timed out

    if index == len(nodes):
        return True

    node = nodes[index]

    forbidden = {
        assignment[n] for n in graph[node] if n in assignment
    }

    for slot in range(num_slots):
        if slot in forbidden:
            continue

        assignment[node] = slot

        # Forward checking
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
                pruned = True
                break

        if not pruned:
            result = _backtrack(nodes, index + 1, graph, assignment,
                                num_slots, counter)
            if result is True:
                return True
            if result is None:
                # Timeout propagated — stop searching this branch
                del assignment[node]
                return None

        del assignment[node]

    return False


def exact_colouring(graph: dict, upper_bound: int) -> tuple[dict, bool]:
    """
    Find the true minimum colouring by trying to beat the upper bound.

    Returns
    -------
    (assignment, exact)
        assignment : { subject_label -> slot_number (0-indexed) }
        exact      : True if the result is proven optimal,
                     False if it is the DSatur result (timeout fallback)
    """
    nodes = sorted(graph.keys(), key=lambda n: len(graph[n]), reverse=True)

    best_assignment = dsatur_colouring(graph)
    exact           = True

    for target in range(upper_bound - 1, 0, -1):
        attempt  = {}
        counter  = [0]

        if nodes:
            attempt[nodes[0]] = 0
            start_index = 1
        else:
            start_index = 0

        result = _backtrack(nodes, start_index, graph, attempt,
                            target, counter)

        if result is True:
            best_assignment = attempt
        elif result is None:
            # Timed out — keep best found so far, mark as not proven exact
            exact = False
            break
        else:
            # Proven impossible — current best_assignment is optimal
            break

    return best_assignment, exact


# ------------------------------------------------
# CLASH REPORT
# ------------------------------------------------
def print_clash_report(exam_tree: ExamTree, exclusions: set = None):
    """
    Print the full exam slot report for every grade.
    Computation is done once per grade and reused for detail + summary.
    """
    if exclusions is None:
        exclusions = set()

    print("=" * 60)
    print("EXAM SLOT REPORT  (exact minimum via backtracking)")
    print("Subjects in the same slot share NO students")
    print("=" * 60)
    print()

    # Compute once per grade, store for summary reuse
    results = {}

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
        assignment, exact = exact_colouring(graph, upper_bound)
        num_slots   = max(assignment.values()) + 1

        results[grade_label] = {
            "student_sets": student_sets,
            "assignment":   assignment,
            "num_slots":    num_slots,
            "exact":        exact,
        }

        # Group subjects by slot
        slots: dict[int, list] = {}
        for subj, slot in assignment.items():
            slots.setdefault(slot, []).append(subj)

        for slot in slots.values():
            slot.sort()

        sorted_slots = sorted(
            slots.values(),
            key=lambda group: sum(len(student_sets[s]) for s in group),
            reverse=True
        )

        exact_note = "" if exact else "  (DSatur fallback — may not be minimum)"
        print(f"{grade_label}  ({num_slots} slots needed){exact_note}")

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

    for grade_label, r in results.items():
        exact_flag = "" if r["exact"] else " *"
        print(f"  {grade_label}: {r['num_slots']} slots{exact_flag}")

    if any(not r["exact"] for r in results.values()):
        print()
        print("  * DSatur fallback used — backtracking timed out.")
        print("    Result is good but may not be the absolute minimum.")

    print()


# ------------------------------------------------
# OPTIONAL TEST NOTE
# ------------------------------------------------
if __name__ == "__main__":
    print("This file is meant to be used from main.py")