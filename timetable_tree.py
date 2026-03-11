"""
timetable_tree.py

Purpose
-------
Read a student timetable file (ST1.csv or ST1.xlsx) and build the timetable tree.

Accepted file formats
---------------------
    ST1.csv
    ST1.xlsx

Tree structure
--------------
TimetableTree
    -> Block
        -> SubBlock
            -> ClassList
                -> StudentList

Example
-------
A
-A1
--AF_BALAY_08
---[1, 6, 8, 9, ...]

Important
---------
The same class (e.g. AF_BALAY_08) can appear in multiple subblocks
across multiple blocks. Each subblock stores its own ClassList object.
The ExamTree is responsible for merging these into a single list per subject.
"""

import re
from pathlib import Path
import pandas as pd
from collections import defaultdict


# ------------------------------------------------
# STUDENT LIST
# ------------------------------------------------
class StudentList:
    """
    Container for student IDs.

    We store students in a set because:
    - duplicates are prevented automatically
    - membership checking is fast
    """

    def __init__(self):
        self.students = set()

    def add_student(self, student_id: int):
        self.students.add(student_id)

    def has_student(self, student_id: int) -> bool:
        return student_id in self.students

    def get_sorted(self):
        """Return a sorted list for display. Underlying data stays a set."""
        return sorted(self.students)

    def __contains__(self, student_id: int):
        return student_id in self.students

    def __len__(self):
        return len(self.students)

    def __str__(self):
        return str(self.get_sorted())


# ------------------------------------------------
# CLASS LIST
# ------------------------------------------------
class ClassList:
    """
    Represents one class in one specific subblock.

    Label format:  SUBJECT_TEACHER_GRADE
    Examples:
        AF_BALAY_08
        MA_ALLEN_10
        DR_VAN_DEN_BERG_12   (multi-word teacher names work fine)

    Note: the same teacher may appear in multiple subblocks.
    Each subblock gets its own ClassList object.
    The ExamTree merges them.
    """

    def __init__(self, label: str):
        self.label = label
        self.student_list = StudentList()

    def add_student(self, student_id: int):
        self.student_list.add_student(student_id)

    def has_student(self, student_id: int) -> bool:
        return self.student_list.has_student(student_id)


# ------------------------------------------------
# SUBBLOCK
# ------------------------------------------------
class SubBlock:
    """
    Represents one timetable slot, e.g. A1, A2, B3, H7.

    A SubBlock contains one ClassList per class running in that slot.
    """

    def __init__(self, name: str):
        self.name = name
        # key = class label, value = ClassList
        self.class_lists = {}

    def get_or_create_class_list(self, class_label: str) -> ClassList:
        if class_label not in self.class_lists:
            self.class_lists[class_label] = ClassList(class_label)
        return self.class_lists[class_label]


# ------------------------------------------------
# BLOCK
# ------------------------------------------------
class Block:
    """
    Represents one timetable block, e.g. A, B, C ... H.

    A Block contains SubBlock objects.
    """

    def __init__(self, name: str):
        self.name = name
        # key = subblock name, value = SubBlock
        self.subblocks = {}

    def get_or_create_subblock(self, subblock_name: str) -> SubBlock:
        if subblock_name not in self.subblocks:
            self.subblocks[subblock_name] = SubBlock(subblock_name)
        return self.subblocks[subblock_name]


# ------------------------------------------------
# ROOT TIMETABLE TREE
# ------------------------------------------------
class TimetableTree:
    """
    Top-level timetable tree.

    Structure:
        TimetableTree
            -> Block
                -> SubBlock
                    -> ClassList
                        -> StudentList
    """

    def __init__(self):
        self.blocks = {}

    def get_or_create_block(self, block_name: str) -> Block:
        if block_name not in self.blocks:
            self.blocks[block_name] = Block(block_name)
        return self.blocks[block_name]

    def add_entry(self, block_name: str, subblock_name: str, class_label: str, student_id: int):
        """
        Insert one student into the correct position in the tree.

        Example:
            block_name    = "A"
            subblock_name = "A1"
            class_label   = "AF_BALAY_08"
            student_id    = 67

        Path:  A -> A1 -> AF_BALAY_08 -> 67
        """
        block    = self.get_or_create_block(block_name)
        subblock = block.get_or_create_subblock(subblock_name)
        cl       = subblock.get_or_create_class_list(class_label)
        cl.add_student(student_id)

    def print_tree(self):
        for block_name in sorted(self.blocks.keys()):
            block = self.blocks[block_name]
            print(block.name)

            sorted_subblocks = sorted(
                block.subblocks.keys(),
                key=lambda name: int(name[1:])
            )

            for subblock_name in sorted_subblocks:
                subblock = block.subblocks[subblock_name]
                print(f"-{subblock.name}")

                for class_label in sorted(subblock.class_lists.keys()):
                    cl = subblock.class_lists[class_label]
                    print(f"--{cl.label}")
                    print(f"---{cl.student_list}")

            print()


# ------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------
def detect_timetable_columns(df: pd.DataFrame):
    """
    Detect columns matching the pattern [A-H] followed by digits.
    Examples: A1, A2, B1, D7, H12

    Returns them sorted block-first, then by subblock number.
    """
    cols = []
    for col in df.columns:
        text = str(col).strip()
        if re.fullmatch(r"[A-H]\d+", text):
            cols.append(text)
    cols.sort(key=lambda x: (x[0], int(x[1:])))
    return cols


# Subject codes that should be merged into another code.
# Any entry found under the key will be stored as the value instead.
SUBJECT_MERGES = {
    "OD": "DR",
}

def build_class_label(raw_label: str, grade) -> str:
    """
    Convert a raw cell value and grade into a standardised class label.

    Applies any subject merges defined in SUBJECT_MERGES before building
    the label, so merged subjects are consolidated at the source.

    Examples:
        "AF BALAY",  8  ->  AF_BALAY_08
        "OD MUNROE", 8  ->  DR_MUNROE_08   (OD merged into DR)
    """
    parts     = str(raw_label).strip().split()
    # Apply subject merge if the subject code is in the merge map
    if parts[0] in SUBJECT_MERGES:
        parts[0] = SUBJECT_MERGES[parts[0]]
    base      = "_".join(parts)
    grade_int = int(grade)
    return f"{base}_{grade_int:02d}"


def check_for_data_warnings(df: pd.DataFrame, timetable_cols: list):
    """
    Scan the data for likely problems and print warnings.

    Flags any class with fewer than 5 students total.
    This catches typos like 'MA BLL' or 'MA BELL' entered as
    'ML BELL' for one student, which create a ghost class.

    Each warning shows:
    - the class label
    - how many students are in it
    - which subblocks it was found in
    - the actual student IDs, so you can check the source data
    """
    # class_label -> { 'students': set, 'subblocks': set }
    class_info = defaultdict(lambda: {"students": set(), "subblocks": set()})

    for _, row in df.iterrows():
        student_id = row.get("Studentid")
        grade      = row.get("Grade")
        if pd.isna(student_id) or pd.isna(grade):
            continue
        for col in timetable_cols:
            val = row[col]
            if pd.isna(val):
                continue
            raw = str(val).strip()
            if not raw:
                continue
            label = build_class_label(raw, grade)
            class_info[label]["students"].add(int(student_id))
            class_info[label]["subblocks"].add(col)

    warnings = []
    for label, info in class_info.items():
        count = len(info["students"])
        if count < 5:
            subblocks  = sorted(info["subblocks"], key=lambda x: (x[0], int(x[1:])))
            student_ids = sorted(info["students"])
            warnings.append(
                f"  '{label}': {count} student(s) "
                f"| subblock(s): {subblocks} "
                f"| student ID(s): {student_ids}"
            )

    if warnings:
        print("DATA WARNINGS (classes with fewer than 5 students - possible typos):")
        for w in sorted(warnings):
            print(w)
        print()


def load_dataframe(file_path) -> pd.DataFrame:
    """
    Load ST1 from either a .csv or .xlsx file.
    Raises a clear error if the format is not recognised.
    """
    path   = Path(file_path)
    suffix = path.suffix.lower()

    if suffix == ".csv":
        # Export uses semicolon delimiters and latin-1 encoding
        return pd.read_csv(file_path, sep=";", encoding="latin-1")
    elif suffix in (".xlsx", ".xls"):
        return pd.read_excel(file_path)
    else:
        raise ValueError(
            f"Unsupported file format: '{suffix}'. "
            "Please provide ST1.csv or ST1.xlsx."
        )


def build_timetable_tree_from_file(st_file_path) -> TimetableTree:
    """
    Read ST1.csv or ST1.xlsx and return a populated TimetableTree.

    Each row in the file represents one student.
    Each timetable column (A1..H7 etc.) holds the class that student
    attends in that slot.

    The same class label may appear in multiple subblocks.
    Each subblock stores its own ClassList independently.
    The ExamTree handles merging.
    """
    df = load_dataframe(st_file_path)

    # Only keep rows with a real student ID
    df = df.dropna(subset=["Studentid"]).copy()

    # Convert Studentid to int safely
    df["Studentid"] = df["Studentid"].astype(float).astype(int)

    timetable_cols = detect_timetable_columns(df)

    if not timetable_cols:
        raise ValueError(
            "No timetable columns found in the file. "
            "Expected columns like A1, B3, H7 etc."
        )

    print(f"Found {len(timetable_cols)} timetable columns: "
          f"{timetable_cols[0]} -> {timetable_cols[-1]}")

    # Warn about possible data entry errors
    check_for_data_warnings(df, timetable_cols)

    tree = TimetableTree()

    for _, row in df.iterrows():
        student_id = int(row["Studentid"])
        grade      = row["Grade"]

        if pd.isna(grade):
            continue

        for subblock_name in timetable_cols:
            subject_value = row[subblock_name]

            if pd.isna(subject_value):
                continue

            raw_label = str(subject_value).strip()
            if not raw_label:
                continue

            block_name  = subblock_name[0]
            class_label = build_class_label(raw_label, grade)

            tree.add_entry(block_name, subblock_name, class_label, student_id)

    return tree


# ------------------------------------------------
# OPTIONAL STANDALONE TEST
# ------------------------------------------------
if __name__ == "__main__":

    for candidate in ["ST1.csv", "ST1.xlsx"]:
        p = Path(candidate)
        if p.exists():
            print(f"Using: {candidate}")
            timetable_tree = build_timetable_tree_from_file(p)
            timetable_tree.print_tree()
            break
    else:
        print("No ST1.csv or ST1.xlsx found in the current folder.")