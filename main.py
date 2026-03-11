"""
main.py

Purpose
-------
1. Find the student timetable file (ST1.csv preferred, ST1.xlsx fallback)
2. Build the TimetableTree
3. Build the ExamTree from the TimetableTree
4. Print both trees

Usage
-----
    python main.py

The script looks for ST1.csv (or ST1.xlsx) in the same folder.
"""

from pathlib import Path
from timetable_tree import build_timetable_tree_from_file
from exam_tree import build_exam_tree_from_timetable_tree
from exam_clash import print_clash_report


def find_student_file(base_folder: Path) -> Path:
    """
    Look for the student timetable file.
    Prefer CSV over xlsx since the new export is CSV.
    """
    for filename in ["ST1.csv", "ST1.xlsx"]:
        candidate = base_folder / filename
        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        "Could not find ST1.csv or ST1.xlsx in the current folder. "
        "Please place the exported file here and try again."
    )


def main():
    base_folder = Path(".")
    st_file = find_student_file(base_folder)
    print(f"Loading: {st_file.name}")
    print()

    # ---- Timetable Tree ----
    timetable_tree = build_timetable_tree_from_file(st_file)

    #print("=" * 60)
    #print("TIMETABLE TREE")
    #print("=" * 60)
    #timetable_tree.print_tree()

    # ---- Exam Tree ----
    exam_tree = build_exam_tree_from_timetable_tree(timetable_tree)

    print("=" * 60)
    print("EXAM TREE")
    print("=" * 60)
    exam_tree.print_tree()

    print("=" * 60)
    print("EXAM CLASH REPORT")
    print("=" * 60)
    print_clash_report(exam_tree)


if __name__ == "__main__":
    main()
