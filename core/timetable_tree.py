"""
timetable_tree.py

Purpose
-------
Read a student timetable file (ST1.xlsx) and build the timetable tree.

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

Cell values skipped during parsing
-----------------------------------
- Empty / NaN cells
- "FREE" cells  (available slots with no class assigned)
"""

import re
from pathlib import Path
import pandas as pd
from collections import defaultdict

# Cell values that represent an empty/available slot — not a real class.
SKIP_VALUES = {"FREE"}


# ------------------------------------------------
# STUDENT LIST
# ------------------------------------------------
class StudentList:
    def __init__(self):
        self.students = set()

    def add_student(self, student_id: int):
        self.students.add(student_id)

    def has_student(self, student_id: int) -> bool:
        return student_id in self.students

    def get_sorted(self):
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
        DR_VAN_DEN_BERG_12
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
    def __init__(self, name: str):
        self.name = name
        self.class_lists = {}

    def get_or_create_class_list(self, class_label: str) -> ClassList:
        if class_label not in self.class_lists:
            self.class_lists[class_label] = ClassList(class_label)
        return self.class_lists[class_label]


# ------------------------------------------------
# BLOCK
# ------------------------------------------------
class Block:
    def __init__(self, name: str):
        self.name = name
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

    def add_entry(self, block_name: str, subblock_name: str,
                  class_label: str, student_id: int):
        block    = self.get_or_create_block(block_name)
        subblock = block.get_or_create_subblock(subblock_name)
        cl       = subblock.get_or_create_class_list(class_label)
        cl.add_student(student_id)

    def print_tree(self):
        for block_name in sorted(self.blocks.keys()):
            block = self.blocks[block_name]
            print(block.name)
            for subblock_name in sorted(block.subblocks,
                                        key=lambda n: int(n[1:])):
                subblock = block.subblocks[subblock_name]
                print(f"-{subblock.name}")
                for class_label in sorted(subblock.class_lists):
                    cl = subblock.class_lists[class_label]
                    print(f"--{cl.label}")
                    print(f"---{cl.student_list}")
            print()


# ------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------

def detect_timetable_columns(df: pd.DataFrame) -> list[str]:
    """
    Detect columns matching the pattern [A-H] followed by digits.
    Returns them sorted block-first, then by subblock number.
    """
    cols = [
        str(col).strip()
        for col in df.columns
        if re.fullmatch(r"[A-H]\d+", str(col).strip())
    ]
    cols.sort(key=lambda x: (x[0], int(x[1:])))
    return cols


# Subject codes that should be merged into another code at parse time.
SUBJECT_MERGES = {
    "OD": "DR",
}


def build_class_label(raw_label: str, grade) -> str:
    """
    Convert a raw cell value and grade into a standardised class label.

    Examples:
        "AF BALAY",  8  ->  AF_BALAY_08
        "OD MUNROE", 8  ->  DR_MUNROE_08   (OD merged into DR)
    """
    parts = str(raw_label).strip().split()
    if parts[0] in SUBJECT_MERGES:
        parts[0] = SUBJECT_MERGES[parts[0]]
    base = "_".join(parts)
    return f"{base}_{int(grade):02d}"


def _should_skip(raw: str) -> bool:
    """
    Return True for cell values that represent empty/free slots.
    Centralised here so the same logic applies in both the parser
    and the data-warning scanner.
    """
    return not raw or raw.upper() in SKIP_VALUES


def check_for_data_warnings(df: pd.DataFrame, timetable_cols: list):
    """
    Scan the data for likely problems and print warnings.
    Flags any class with fewer than 5 students — catches typos that
    create ghost classes (e.g. 'MA BLL' for one student).
    FREE cells are ignored.
    """
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
            if _should_skip(raw):
                continue
            label = build_class_label(raw, grade)
            class_info[label]["students"].add(int(student_id))
            class_info[label]["subblocks"].add(col)

    warnings = []
    for label, info in class_info.items():
        count = len(info["students"])
        if count < 5:
            subblocks   = sorted(info["subblocks"],
                                 key=lambda x: (x[0], int(x[1:])))
            student_ids = sorted(info["students"])
            warnings.append(
                f"  '{label}': {count} student(s) "
                f"| subblock(s): {subblocks} "
                f"| student ID(s): {student_ids}"
            )

    if warnings:
        print("DATA WARNINGS (classes with fewer than 5 students):")
        for w in sorted(warnings):
            print(w)
        print()


def load_dataframe(file_path) -> pd.DataFrame:
    path   = Path(file_path)
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        return pd.read_excel(file_path)
    raise ValueError(
        f"Unsupported file format: '{suffix}'. Please provide ST1.xlsx."
    )


def build_timetable_tree_from_file(st_file_path) -> TimetableTree:
    """
    Read ST1.xlsx and return a populated TimetableTree.

    Each row represents one student.
    Each timetable column (A1..H7 etc.) holds the class for that slot.
    FREE cells and empty cells are silently skipped.
    """
    df = load_dataframe(st_file_path)
    df = df.dropna(subset=["Studentid"]).copy()
    df["Studentid"] = df["Studentid"].astype(float).astype(int)

    timetable_cols = detect_timetable_columns(df)
    if not timetable_cols:
        raise ValueError(
            "No timetable columns found. Expected columns like A1, B3, H7."
        )

    print(f"Found {len(timetable_cols)} timetable columns: "
          f"{timetable_cols[0]} -> {timetable_cols[-1]}")

    check_for_data_warnings(df, timetable_cols)

    tree = TimetableTree()

    for _, row in df.iterrows():
        student_id = int(row["Studentid"])
        grade      = row["Grade"]
        if pd.isna(grade):
            continue

        for subblock_name in timetable_cols:
            val = row[subblock_name]
            if pd.isna(val):
                continue
            raw = str(val).strip()
            if _should_skip(raw):
                continue

            block_name  = subblock_name[0]
            class_label = build_class_label(raw, grade)
            tree.add_entry(block_name, subblock_name, class_label, student_id)

    return tree


# ------------------------------------------------
# OPTIONAL STANDALONE TEST
# ------------------------------------------------
if __name__ == "__main__":
    for candidate in ["data/ST1.xlsx"]:
        p = Path(candidate)
        if p.exists():
            print(f"Using: {candidate}")
            build_timetable_tree_from_file(p).print_tree()
            break
    else:
        print("data/ST1.xlsx not found.")