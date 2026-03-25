"""
timetable_tree.py

Purpose
-------
Read a student timetable file (ST1.xlsx) and build the timetable tree.

CHANGELOG (bug-fix edition)
----------------------------
BUG 2 FIXED — false timetable column detection
    Old regex: [A-H]\\d+
    Problem:   This matched ANY letter A-H followed by ANY number of digits.
               ST1.xlsx contains exam columns F25, F26 ... F67 (used by the
               exam scheduler).  All 43 of these were incorrectly treated as
               timetable subblocks, inflating the column list and adding
               garbage data to every student's tree entry for those slots.
    Fix:       ^[A-H][1-9]$  — letter A-H + exactly one digit 1-9.
               Real timetable subblocks are A1-A7 through H1-H7 (single
               digit suffixes only).  Exam columns like F25 have two-digit
               suffixes and are now excluded cleanly.

Tree structure (unchanged)
--------------------------
TimetableTree
    -> Block
        -> SubBlock
            -> ClassList
                -> StudentList
"""

import re
from pathlib import Path
import pandas as pd
from collections import defaultdict


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
    def __init__(self):
        self.blocks = {}
        # Maps student_id -> "Firstname Lastname" (populated from Excel Name/Surname columns)
        self.student_names: dict[int, str] = {}

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
# SERIALISATION
# ------------------------------------------------

def timetable_tree_to_dict(tree: "TimetableTree") -> dict:
    """Serialise a TimetableTree to a plain dict suitable for json.dump."""
    out: dict = {}
    for block_name, block in tree.blocks.items():
        out[block_name] = {}
        for sb_name, subblock in block.subblocks.items():
            out[block_name][sb_name] = {}
            for cl_label, cl in subblock.class_lists.items():
                out[block_name][sb_name][cl_label] = sorted(
                    cl.student_list.students)
    return {
        "_blocks": out,
        "_names": {str(k): v for k, v in tree.student_names.items()},
    }


def timetable_tree_from_dict(data: dict) -> "TimetableTree":
    """Reconstruct a TimetableTree from a dict produced by timetable_tree_to_dict."""
    tree = TimetableTree()
    # Support both new format (with _blocks/_names keys) and legacy format
    if "_blocks" in data:
        blocks_data = data["_blocks"]
        tree.student_names = {int(k): v for k, v in data.get("_names", {}).items()}
    else:
        blocks_data = data
    for block_name, subblocks in blocks_data.items():
        for sb_name, class_lists in subblocks.items():
            for cl_label, students in class_lists.items():
                for sid in students:
                    tree.add_entry(block_name, sb_name, cl_label, int(sid))
    return tree


# ------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------

# BUG 2 FIX: old pattern [A-H]\\d+ matched F25, F26 ... F67.
# New pattern matches only single-digit subblocks (A1-H9).
# Real timetable columns are A1-H7; two-digit exam columns are excluded.
_TIMETABLE_COL_RE = re.compile(r"^[A-H][1-9]$")

def detect_timetable_columns(df: pd.DataFrame) -> list[str]:
    """
    Detect columns matching the pattern [A-H] followed by ONE digit 1-9.
    e.g.  A1, A2, B1, D7, H7  —  but NOT F25, F30, F67 (exam columns).

    Sorted block-first, then by subblock number.
    """
    cols = [
        str(col).strip()
        for col in df.columns
        if _TIMETABLE_COL_RE.match(str(col).strip())
    ]
    cols.sort(key=lambda x: (x[0], int(x[1:])))
    return cols


# Subject codes that should be merged into another code at load time.
SUBJECT_MERGES = {
    "OD": "DR",
}

def build_class_label(raw_label: str, grade) -> str:
    """
    Convert a raw cell value and grade into a standardised class label.
    Applies subject merges before building the label.

    Examples:
        "AF BALAY",  8  ->  AF_BALAY_08
        "OD MUNROE", 8  ->  DR_MUNROE_08   (OD merged into DR)
    """
    parts     = str(raw_label).strip().split()
    if parts[0] in SUBJECT_MERGES:
        parts[0] = SUBJECT_MERGES[parts[0]]
    base      = "_".join(parts)
    grade_int = int(grade)
    return f"{base}_{grade_int:02d}"


def check_for_data_warnings(df: pd.DataFrame, timetable_cols: list):
    """
    Scan the data for likely problems and print warnings.
    Flags any class with fewer than 5 students total.
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
            if not raw:
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

    The fixed column detection regex (r"^[A-H][1-9]$") means the 43
    spurious F25-F67 exam columns are silently ignored here — they will
    never appear in the tree and will never pollute a student's schedule.
    """
    df = load_dataframe(st_file_path)
    df = df.dropna(subset=["Studentid"]).copy()
    df["Studentid"] = df["Studentid"].astype(float).astype(int)

    timetable_cols = detect_timetable_columns(df)

    if not timetable_cols:
        raise ValueError(
            "No timetable columns found in the file. "
            "Expected columns like A1, B3, H7 (single-digit suffix)."
        )

    print(f"Found {len(timetable_cols)} timetable columns: "
          f"{timetable_cols[0]} -> {timetable_cols[-1]}")

    check_for_data_warnings(df, timetable_cols)

    tree = TimetableTree()

    # Build student_names lookup from Name/Surname columns if present
    has_name    = "Name"    in df.columns
    has_surname = "Surname" in df.columns
    for _, row in df.iterrows():
        sid = int(row["Studentid"])
        parts = []
        if has_name and not pd.isna(row.get("Name")):
            parts.append(str(row["Name"]).strip())
        if has_surname and not pd.isna(row.get("Surname")):
            parts.append(str(row["Surname"]).strip())
        if parts:
            tree.student_names[sid] = " ".join(parts)

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

    candidates = [
        "data/ST1.xlsx",       # run from project root  (python core/timetable_tree.py)
        "../data/ST1.xlsx",    # run from core/          (python timetable_tree.py)
        "ST1.xlsx",            # run from data/
    ]
    for candidate in candidates:
        p = Path(candidate)
        if p.exists():
            print(f"Using: {candidate}")
            timetable_tree = build_timetable_tree_from_file(p)
            timetable_tree.print_tree()
            break
    else:
        print(f"ST1.xlsx not found. Tried: {candidates}")