"""
verify_timetable.py
-------------------
Verifies whether a completed TimetableTree is legal.

Hard constraints checked
------------------------
1. Student double-booking  — a student appears in more than one class
                             within the same subblock (same time slot).
2. Teacher double-booking  — a teacher appears in more than one class
                             within the same subblock.

Usage
-----
    from verify_timetable import verify_timetable

    is_legal = verify_timetable(timetable_tree)            # silent
    is_legal = verify_timetable(timetable_tree, verbose=True)  # prints full report

    # Or call the report directly (always prints):
    from verify_timetable import print_clash_report
    print_clash_report(timetable_tree)

Returns
-------
    True   — no violations found; timetable is legal
    False  — one or more violations found
"""

from __future__ import annotations
from collections import defaultdict
from timetable_tree import TimetableTree


# ---------------------------------------------------------------------------
# Label helpers
# ---------------------------------------------------------------------------

def extract_teacher(class_label: str) -> str:
    """
    Pull the teacher identifier out of a class label.

    Label format:  SUBJECT_TEACHER_GRADE
    Examples:
        AF_BALAY_08          -> BALAY
        MA_ALLEN_10          -> ALLEN
        DR_VAN_DEN_BERG_12   -> VAN_DEN_BERG
    """
    parts = class_label.split("_")
    # parts[0]  = subject code
    # parts[-1] = grade (two-digit number)
    # everything in between = teacher
    return "_".join(parts[1:-1])


def extract_subject(class_label: str) -> str:
    return class_label.split("_")[0]


def extract_grade(class_label: str) -> str:
    return class_label.split("_")[-1]


# ---------------------------------------------------------------------------
# Core clash detection
# ---------------------------------------------------------------------------

def _find_clashes(tree: TimetableTree) -> tuple[list[dict], list[dict]]:
    """
    Walk every subblock in the tree and collect all violations.

    Returns
    -------
    student_clashes : list of dicts, one per (subblock, student) pair
                      that appears in more than one class.
    teacher_clashes : list of dicts, one per (subblock, teacher) pair
                      that appears in more than one class.
    """
    student_clashes: list[dict] = []
    teacher_clashes: list[dict] = []

    for block_name in sorted(tree.blocks):
        block = tree.blocks[block_name]

        for subblock_name in sorted(block.subblocks,
                                    key=lambda n: (n[0], int(n[1:]))):
            subblock = block.subblocks[subblock_name]

            # Build lookup maps for this subblock
            # student_id  -> list of class labels they appear in
            # teacher_key -> list of class labels they teach
            student_to_classes: dict[int, list[str]] = defaultdict(list)
            teacher_to_classes: dict[str, list[str]] = defaultdict(list)

            for class_label, class_list in subblock.class_lists.items():
                teacher = extract_teacher(class_label)
                teacher_to_classes[teacher].append(class_label)

                for student_id in class_list.student_list.students:
                    student_to_classes[student_id].append(class_label)

            # --- student violations ---
            for student_id, classes in student_to_classes.items():
                if len(classes) > 1:
                    student_clashes.append({
                        "subblock": subblock_name,
                        "student":  student_id,
                        "classes":  sorted(classes),
                    })

            # --- teacher violations ---
            for teacher, classes in teacher_to_classes.items():
                if len(classes) > 1:
                    teacher_clashes.append({
                        "subblock": subblock_name,
                        "teacher":  teacher,
                        "classes":  sorted(classes),
                    })

    return student_clashes, teacher_clashes


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def verify_timetable(tree: TimetableTree, verbose: bool = False) -> bool:
    """
    Check whether a TimetableTree satisfies all hard constraints.

    Parameters
    ----------
    tree    : a fully populated TimetableTree
    verbose : if True, print a full clash report when violations exist

    Returns
    -------
    True  — timetable is legal (no violations)
    False — one or more violations found
    """
    student_clashes, teacher_clashes = _find_clashes(tree)
    is_legal = not student_clashes and not teacher_clashes

    if verbose and not is_legal:
        print_clash_report(tree)
    elif verbose:
        print("Timetable verification PASSED — no violations found.")

    return is_legal


def print_clash_report(tree: TimetableTree) -> None:
    """
    Run the legality check and print a full human-readable clash report.
    Always prints, regardless of the result.
    """
    student_clashes, teacher_clashes = _find_clashes(tree)
    is_legal = not student_clashes and not teacher_clashes

    _print_header(is_legal, student_clashes, teacher_clashes)

    if student_clashes:
        _print_student_clashes(student_clashes)

    if teacher_clashes:
        _print_teacher_clashes(teacher_clashes)

    if is_legal:
        print("  No violations detected.\n")

    _print_summary(student_clashes, teacher_clashes)


# ---------------------------------------------------------------------------
# Report formatting helpers
# ---------------------------------------------------------------------------

def _print_header(is_legal: bool, student_clashes: list, teacher_clashes: list):
    status = "PASSED ✓" if is_legal else "FAILED ✗"
    print()
    print("=" * 60)
    print(f"  TIMETABLE LEGALITY CHECK — {status}")
    print("=" * 60)
    if not is_legal:
        print(f"  Student double-bookings : {len(student_clashes)}")
        print(f"  Teacher double-bookings : {len(teacher_clashes)}")
    print()


def _print_student_clashes(clashes: list[dict]):
    # Group by subblock for a cleaner report
    by_subblock: dict[str, list[dict]] = defaultdict(list)
    for c in clashes:
        by_subblock[c["subblock"]].append(c)

    print("  STUDENT DOUBLE-BOOKINGS")
    print("  " + "-" * 56)

    for subblock in sorted(by_subblock, key=lambda n: (n[0], int(n[1:]))):
        entries = by_subblock[subblock]
        print(f"\n  Subblock {subblock}  ({len(entries)} student(s) affected)")
        for entry in sorted(entries, key=lambda e: e["student"]):
            classes_str = "  vs  ".join(entry["classes"])
            print(f"    Student {entry['student']:>6} :  {classes_str}")
    print()


def _print_teacher_clashes(clashes: list[dict]):
    by_subblock: dict[str, list[dict]] = defaultdict(list)
    for c in clashes:
        by_subblock[c["subblock"]].append(c)

    print("  TEACHER DOUBLE-BOOKINGS")
    print("  " + "-" * 56)

    for subblock in sorted(by_subblock, key=lambda n: (n[0], int(n[1:]))):
        entries = by_subblock[subblock]
        print(f"\n  Subblock {subblock}  ({len(entries)} teacher(s) affected)")
        for entry in sorted(entries, key=lambda e: e["teacher"]):
            classes_str = "  vs  ".join(entry["classes"])
            print(f"    {entry['teacher']:<20} :  {classes_str}")
    print()


def _print_summary(student_clashes: list, teacher_clashes: list):
    print("=" * 60)
    total = len(student_clashes) + len(teacher_clashes)
    if total == 0:
        print("  Total violations: 0 — timetable is schedulable.")
    else:
        print(f"  Total violations: {total}  "
              f"({len(student_clashes)} student, {len(teacher_clashes)} teacher)")
    print("=" * 60)
    print()


# ---------------------------------------------------------------------------
# Standalone usage
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from pathlib import Path
    from timetable_tree import build_timetable_tree_from_file

    for candidate in ["ST1.csv", "data/ST1.xlsx"]:
        p = Path(candidate)
        if p.exists():
            print(f"Loading: {candidate}")
            tree = build_timetable_tree_from_file(p)
            print_clash_report(tree)
            break
    else:
        print("No ST1.csv or ST1.xlsx found in the current folder.")