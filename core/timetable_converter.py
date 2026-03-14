"""
timetable_converter.py
----------------------
Converts a read-only TimetableTree into a mutable BlockTree.

This is the bridge between Pipeline B's read side and its write/optimise side.
Once converted, the BlockTree can be:
    - Scored by cost_function.evaluate()         (read from TimetableTree directly)
    - Mutated by the SA optimiser                (swap_assignments, move_assignment)
    - Exported back to ST1.xlsx                  (block_exporter.export_to_xlsx)

Conversion logic
----------------
TimetableTree leaf:
    ClassList("AF_BALAY_08")
        StudentList: {361, 370, 375}

BlockTree leaf:
    SlotAssignment(
        subject_code = "AF"
        teacher_code = "BALAY"
        grade        = "Gr_08"
        student_ids  = {361, 370, 375}
    )

The label is parsed back into components. Grade is converted from the
TimetableTree format ("08") to the BlockTree format ("Gr_08").

The full Block → SubBlock structure is preserved exactly.
"""

from core.timetable_tree import TimetableTree
from core.block_tree      import BlockTree


def _parse_label(class_label: str) -> tuple[str, str, str]:
    """
    Split a class label into (subject_code, teacher_code, grade).

    Label format: SUBJECT_TEACHER_GRADE
    Examples:
        "AF_BALAY_08"        ->  ("AF",  "BALAY",         "08")
        "MA_ALLEN_10"        ->  ("MA",  "ALLEN",         "10")
        "DR_VAN_DEN_BERG_12" ->  ("DR",  "VAN_DEN_BERG",  "12")
    """
    parts        = class_label.split("_")
    subject_code = parts[0]
    grade        = parts[-1]
    teacher_code = "_".join(parts[1:-1])
    return subject_code, teacher_code, grade


def _to_block_tree_grade(grade: str) -> str:
    """
    Convert TimetableTree grade format to BlockTree grade format.
        "08"  ->  "Gr_08"
        "12"  ->  "Gr_12"
    """
    return f"Gr_{grade}"


def timetable_tree_to_block_tree(tt: TimetableTree) -> BlockTree:
    """
    Convert a TimetableTree to a BlockTree.

    Parameters
    ----------
    tt : a fully populated TimetableTree (from build_timetable_tree_from_file)

    Returns
    -------
    BlockTree with the same structure and student assignments.
    BlockTree.validate() should return an empty list if the source
    timetable has no clashes.
    """
    # Initialise BlockTree with only the blocks that actually exist in the
    # timetable — do not assume A-H or 7 subblocks per block.
    block_names = sorted(tt.blocks.keys())
    bt = BlockTree(block_names=block_names, subblocks_per_block=0)

    for block_name, block in tt.blocks.items():
        for subblock_name, subblock in block.subblocks.items():
            for class_label, class_list in subblock.class_lists.items():

                subject_code, teacher_code, grade_raw = _parse_label(class_label)
                grade       = _to_block_tree_grade(grade_raw)
                student_ids = set(class_list.student_list.students)

                bt.assign(
                    subblock_name = subblock_name,
                    subject_code  = subject_code,
                    grade         = grade,
                    teacher_code  = teacher_code,
                    student_ids   = student_ids,
                )

    return bt


# ---------------------------------------------------------------------------
# Standalone test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from pathlib import Path
    from core.timetable_tree import build_timetable_tree_from_file

    st_file = Path("E:\TimePyBling\data\ST1.xlsx")
    if not st_file.exists():
        print("data/ST1.xlsx not found.")
        raise SystemExit(1)

    print(f"Loading: {st_file}")
    tt = build_timetable_tree_from_file(st_file)
    bt = timetable_tree_to_block_tree(tt)

    errors = bt.validate()
    if errors:
        print(f"\n{len(errors)} validation error(s):")
        for e in errors:
            print(f"  {e}")
    else:
        print("\nBlockTree validated — no clashes.")

    # Spot check: pick the first student and print their schedule
    all_ids = bt.all_student_ids()
    if all_ids:
        sample_id = sorted(all_ids)[0]
        schedule  = bt.find_student_schedule(sample_id)
        print(f"\nSchedule for student {sample_id}:")
        for sb_name, assignment in sorted(schedule.items()):
            print(f"  {sb_name}  {assignment.label}")