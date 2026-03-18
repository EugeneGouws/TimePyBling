"""
verify_timetable.py
-------------------
Detects hard constraint violations in a loaded TimetableTree.

Checks
------
  Student double-bookings — a student appears in more than one class
                            within the same subblock.

Usage
-----
    from reader.verify_timetable import find_student_clashes

    clashes = find_student_clashes(timetable_tree)
    for c in clashes:
        print(c["subblock"], c["student"], c["classes"])
"""

from __future__ import annotations
from core.timetable_tree import TimetableTree


def find_student_clashes(tree: TimetableTree) -> list[dict]:
    """
    Scan the timetable tree for student double-bookings.

    A clash occurs when a student appears in two or more class lists
    in the same subblock.

    Returns
    -------
    List of clash dicts, each with keys:
        subblock : str        — e.g. "A3"
        student  : int        — student ID
        classes  : list[str]  — the class labels the student is double-booked in
    """
    clashes = []

    for block in tree.blocks.values():
        for sb_name, subblock in block.subblocks.items():
            # Map student_id -> list of class labels they appear in
            student_classes: dict[int, list[str]] = {}

            for class_label, class_list in subblock.class_lists.items():
                for student_id in class_list.student_list.students:
                    student_classes.setdefault(student_id, []).append(class_label)

            for student_id, labels in student_classes.items():
                if len(labels) > 1:
                    clashes.append({
                        "subblock": sb_name,
                        "student":  student_id,
                        "classes":  labels,
                    })

    return clashes


# ---------------------------------------------------------------------------
# Backward-compatible shim
# ---------------------------------------------------------------------------

def _find_clashes(tree: TimetableTree) -> tuple[list[dict], list[dict]]:
    """
    Legacy two-tuple interface kept for any callers that expect
    (student_clashes, teacher_clashes).

    Teacher checking has been removed from this tool. The second
    element of the tuple is always an empty list.
    """
    return find_student_clashes(tree), []


# ---------------------------------------------------------------------------
# Standalone runner
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from pathlib import Path
    from core.timetable_tree import build_timetable_tree_from_file

    st_file = Path("data/ST1.xlsx")
    if not st_file.exists():
        print("data/ST1.xlsx not found.")
        raise SystemExit(1)

    print(f"Loading: {st_file}")
    tree    = build_timetable_tree_from_file(st_file)
    clashes = find_student_clashes(tree)

    if not clashes:
        print("PASS ✓  — no student clashes found.")
    else:
        print(f"FAIL ✗  — {len(clashes)} student clash(es):")
        for c in sorted(clashes, key=lambda x: (x["subblock"], x["student"])):
            print(f"  {c['subblock']}  student {c['student']:>6}:  "
                  f"{' vs '.join(c['classes'])}")