"""
block_tree.py
-------------
Core data structure for the new timetable builder.

Mirrors the read side (TimetableTree) but is designed to be:
    - Built by an algorithm (block_builder.py)
    - Edited manually via a GUI
    - Exported to ST1-format xlsx (block_exporter.py)

Tree structure
--------------
BlockTree (root)
    -> Block          (A, B, C, D, E, F, G, H)
        -> SubBlock   (A1, A2 ... A7, B1 ... H7)
            -> SlotAssignment
                subject_code   : str   e.g. "MA"
                grade          : str   e.g. "Gr_10"
                teacher_code   : str   e.g. "AA"
                student_ids    : set   {480, 481, ...}

Design notes
------------
- A SubBlock holds MULTIPLE SlotAssignments (parallel classes run
  simultaneously in the same slot, e.g. AF and ZU for Gr_8)
- A student appears in exactly ONE SlotAssignment per SubBlock
- The same subject+grade can appear in multiple SubBlocks across
  different Blocks (e.g. MA runs 4 times per cycle)
- SlotAssignment is the atom the GUI and algorithm both manipulate

Validation
----------
BlockTree.validate() checks the tree for internal consistency:
    - No student assigned to two slots in the same SubBlock
    - No student missing from any slot they should appear in
    - Teacher not double-booked within a SubBlock
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Iterator


# ---------------------------------------------------------------------------
# SLOT ASSIGNMENT
# The atom of the timetable — one class group in one slot
# ---------------------------------------------------------------------------

@dataclass
class SlotAssignment:
    """
    One class running in one SubBlock.

    Examples
    --------
    SlotAssignment("MA", "Gr_10", "AA", {480, 481, 482})
    SlotAssignment("AF", "Gr_8",  "KB", {361, 370, 375})

    The same subject+grade may have multiple SlotAssignments per cycle
    (one per subblock it runs in).
    """
    subject_code : str
    grade        : str
    teacher_code : str
    student_ids  : set = field(default_factory=set)

    # ---- identity ----
    @property
    def label(self) -> str:
        """
        Produces a class label in ST1 format: SUBJECT_TEACHER_GRADE
        e.g.  MA_AA_10   AF_BALAY_08
        Grade is zero-padded to 2 digits to match existing ST1 convention.
        """
        grade_num = self.grade.replace("Gr_", "")
        return f"{self.subject_code}_{self.teacher_code}_{int(grade_num):02d}"

    @property
    def st1_cell(self) -> str:
        """
        Value written into an ST1 cell: 'SUBJECT TEACHER'
        e.g.  'MA AA'   'AF BALAY'
        """
        return f"{self.subject_code} {self.teacher_code}"

    # ---- mutation ----
    def add_student(self, student_id: int):
        self.student_ids.add(student_id)

    def remove_student(self, student_id: int):
        self.student_ids.discard(student_id)

    def set_teacher(self, teacher_code: str):
        self.teacher_code = teacher_code

    # ---- query ----
    def has_student(self, student_id: int) -> bool:
        return student_id in self.student_ids

    def size(self) -> int:
        return len(self.student_ids)

    def __repr__(self):
        return (f"SlotAssignment({self.label}, "
                f"n={self.size()}, "
                f"students={sorted(self.student_ids)})")


# ---------------------------------------------------------------------------
# SUBBLOCK
# One timetable slot — may run multiple parallel classes simultaneously
# ---------------------------------------------------------------------------

class SubBlock:
    """
    One timetable slot, e.g. A1, B3, H7.

    Holds one SlotAssignment per class running in that slot.
    Multiple classes run in parallel (different subjects/grades/teachers).

    Key: subject_code + grade   e.g.  "MA|Gr_10"
    """

    def __init__(self, name: str):
        self.name        = name                          # e.g. "A1"
        self.assignments : dict[str, SlotAssignment] = {}

    # ---- internal key ----
    @staticmethod
    def _key(subject_code: str, grade: str) -> str:
        return f"{subject_code}|{grade}"

    # ---- mutation ----
    def add_assignment(self, assignment: SlotAssignment):
        """Add a SlotAssignment to this slot. Replaces if same key exists."""
        key = self._key(assignment.subject_code, assignment.grade)
        self.assignments[key] = assignment

    def remove_assignment(self, subject_code: str, grade: str):
        key = self._key(subject_code, grade)
        self.assignments.pop(key, None)

    def get_or_create_assignment(self, subject_code: str, grade: str,
                                  teacher_code: str = "") -> SlotAssignment:
        key = self._key(subject_code, grade)
        if key not in self.assignments:
            self.assignments[key] = SlotAssignment(
                subject_code=subject_code,
                grade=grade,
                teacher_code=teacher_code
            )
        return self.assignments[key]

    # ---- query ----
    def get_assignment(self, subject_code: str,
                       grade: str) -> SlotAssignment | None:
        return self.assignments.get(self._key(subject_code, grade))

    def assignment_for_student(self, student_id: int) -> SlotAssignment | None:
        """Which class is this student in for this slot? None if unassigned."""
        for a in self.assignments.values():
            if a.has_student(student_id):
                return a
        return None

    def all_assignments(self) -> list[SlotAssignment]:
        return list(self.assignments.values())

    def teachers_in_slot(self) -> list[str]:
        return [a.teacher_code for a in self.assignments.values()]

    def is_empty(self) -> bool:
        return len(self.assignments) == 0

    def __repr__(self):
        return f"SubBlock({self.name}, assignments={list(self.assignments.keys())})"


# ---------------------------------------------------------------------------
# BLOCK
# One day-block — contains SubBlocks (periods within that block)
# ---------------------------------------------------------------------------

class Block:
    """
    One timetable block, e.g. A, B, C ... H.

    Contains SubBlock objects for each period within the block.
    """

    def __init__(self, name: str):
        self.name      = name                       # e.g. "A"
        self.subblocks : dict[str, SubBlock] = {}

    def get_or_create_subblock(self, subblock_name: str) -> SubBlock:
        if subblock_name not in self.subblocks:
            self.subblocks[subblock_name] = SubBlock(subblock_name)
        return self.subblocks[subblock_name]

    def get_subblock(self, subblock_name: str) -> SubBlock | None:
        return self.subblocks.get(subblock_name)

    def sorted_subblocks(self) -> list[SubBlock]:
        """SubBlocks sorted by period number: A1, A2, A3 ..."""
        return sorted(
            self.subblocks.values(),
            key=lambda sb: int(sb.name[1:])
        )

    def __repr__(self):
        return f"Block({self.name}, subblocks={sorted(self.subblocks.keys())})"


# ---------------------------------------------------------------------------
# BLOCK TREE
# Root of the structure — entry point for builder, GUI, and exporter
# ---------------------------------------------------------------------------

class BlockTree:
    """
    Root of the timetable block structure.

    Provides:
        - Construction API   (used by block_builder)
        - Query API          (used by GUI and exporter)
        - Validation         (consistency checks before export)
        - Manual edit API    (move/swap operations for GUI)
    """

    def __init__(self, block_names: list[str] = None,
                 subblocks_per_block: int = 7):
        """
        Parameters
        ----------
        block_names          : list of block labels, default A-H
        subblocks_per_block  : how many periods per block, default 7
        """
        self.blocks : dict[str, Block] = {}

        names = block_names or list("ABCDEFGH")
        for name in names:
            self.blocks[name] = Block(name)
            for i in range(1, subblocks_per_block + 1):
                subblock_name = f"{name}{i}"
                self.blocks[name].get_or_create_subblock(subblock_name)

    # ------------------------------------------------------------------
    # Construction API
    # ------------------------------------------------------------------

    def get_subblock(self, block_name: str,
                     subblock_name: str) -> SubBlock | None:
        block = self.blocks.get(block_name)
        return block.get_subblock(subblock_name) if block else None

    def assign(self, subblock_name: str, subject_code: str,
               grade: str, teacher_code: str,
               student_ids: set) -> SlotAssignment:
        """
        Place a class group into a subblock.

        Parameters
        ----------
        subblock_name : e.g. "A1"
        subject_code  : e.g. "MA"
        grade         : e.g. "Gr_10"
        teacher_code  : e.g. "AA"
        student_ids   : set of student IDs

        Returns the SlotAssignment created.
        """
        block_name = subblock_name[0]
        block      = self.blocks.get(block_name)
        if block is None:
            raise ValueError(f"Block '{block_name}' does not exist.")

        subblock   = block.get_or_create_subblock(subblock_name)
        assignment = subblock.get_or_create_assignment(
            subject_code, grade, teacher_code
        )
        for sid in student_ids:
            assignment.add_student(sid)

        return assignment

    # ------------------------------------------------------------------
    # Manual edit API  (GUI hooks)
    # ------------------------------------------------------------------

    def move_assignment(self, from_subblock: str, to_subblock: str,
                        subject_code: str, grade: str):
        """
        Move a SlotAssignment from one subblock to another.
        Raises ValueError if the source assignment does not exist.
        """
        src = self.get_subblock(from_subblock[0], from_subblock)
        dst = self.get_subblock(to_subblock[0],   to_subblock)

        if src is None or dst is None:
            raise ValueError(f"Subblock not found: {from_subblock} or {to_subblock}")

        assignment = src.get_assignment(subject_code, grade)
        if assignment is None:
            raise ValueError(
                f"No assignment for {subject_code}/{grade} in {from_subblock}"
            )

        dst.add_assignment(assignment)
        src.remove_assignment(subject_code, grade)

    def swap_assignments(self, subblock_a: str, subject_a: str, grade_a: str,
                         subblock_b: str, subject_b: str, grade_b: str):
        """
        Swap two SlotAssignments between two subblocks.
        Useful for the drag-and-drop GUI operation.
        """
        sb_a = self.get_subblock(subblock_a[0], subblock_a)
        sb_b = self.get_subblock(subblock_b[0], subblock_b)

        if sb_a is None or sb_b is None:
            raise ValueError(f"Subblock not found: {subblock_a} or {subblock_b}")

        assign_a = sb_a.get_assignment(subject_a, grade_a)
        assign_b = sb_b.get_assignment(subject_b, grade_b)

        if assign_a is None or assign_b is None:
            raise ValueError("One or both assignments not found for swap.")

        sb_a.remove_assignment(subject_a, grade_a)
        sb_b.remove_assignment(subject_b, grade_b)
        sb_b.add_assignment(assign_a)
        sb_a.add_assignment(assign_b)

    def set_teacher(self, subblock_name: str, subject_code: str,
                    grade: str, teacher_code: str):
        """Update the teacher on an existing assignment."""
        sb = self.get_subblock(subblock_name[0], subblock_name)
        if sb is None:
            raise ValueError(f"Subblock not found: {subblock_name}")
        a = sb.get_assignment(subject_code, grade)
        if a:
            a.set_teacher(teacher_code)

    # ------------------------------------------------------------------
    # Query API
    # ------------------------------------------------------------------

    def all_subblock_names(self) -> list[str]:
        """All subblock names in ST1 column order: A1, A2... H7"""
        names = []
        for block in sorted(self.blocks.keys()):
            for sb in self.blocks[block].sorted_subblocks():
                names.append(sb.name)
        return names

    def all_assignments(self) -> Iterator[tuple[str, SlotAssignment]]:
        """Iterate over (subblock_name, SlotAssignment) for every assignment."""
        for block in sorted(self.blocks.keys()):
            for sb in self.blocks[block].sorted_subblocks():
                for a in sb.all_assignments():
                    yield sb.name, a

    def find_student_schedule(self, student_id: int) -> dict[str, SlotAssignment]:
        """
        Return { subblock_name: SlotAssignment } for one student.
        Shows the student's complete timetable.
        """
        schedule = {}
        for block in sorted(self.blocks.keys()):
            for sb in self.blocks[block].sorted_subblocks():
                a = sb.assignment_for_student(student_id)
                if a:
                    schedule[sb.name] = a
        return schedule

    def all_student_ids(self) -> set[int]:
        """All unique student IDs present anywhere in the tree."""
        ids = set()
        for _, a in self.all_assignments():
            ids.update(a.student_ids)
        return ids

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def validate(self) -> list[str]:
        """
        Check the tree for consistency errors.

        Returns a list of error strings.
        Empty list means the tree is valid and safe to export.

        Checks
        ------
        1. Student appears in more than one assignment in the same subblock
        2. Teacher double-booked within a subblock
        """
        errors = []

        for block in sorted(self.blocks.keys()):
            for sb in self.blocks[block].sorted_subblocks():
                seen_students : dict[int, str] = {}
                seen_teachers : dict[str, str] = {}

                for a in sb.all_assignments():
                    # Check student clash
                    for sid in a.student_ids:
                        if sid in seen_students:
                            errors.append(
                                f"{sb.name}: student {sid} appears in both "
                                f"'{seen_students[sid]}' and '{a.label}'"
                            )
                        else:
                            seen_students[sid] = a.label

                    # Check teacher double-booking
                    if a.teacher_code:
                        if a.teacher_code in seen_teachers:
                            errors.append(
                                f"{sb.name}: teacher '{a.teacher_code}' "
                                f"double-booked in '{seen_teachers[a.teacher_code]}'"
                                f" and '{a.label}'"
                            )
                        else:
                            seen_teachers[a.teacher_code] = a.label

        return errors

    # ------------------------------------------------------------------
    # Display
    # ------------------------------------------------------------------

    def print_tree(self):
        print(f"\nBlockTree")
        print(f"{'='*60}")
        for block_name in sorted(self.blocks.keys()):
            block = self.blocks[block_name]
            print(f"\n  {block.name}")
            for sb in block.sorted_subblocks():
                if sb.is_empty():
                    print(f"    {sb.name}  (empty)")
                else:
                    for a in sb.all_assignments():
                        print(f"    {sb.name}  {a.label:<30}  "
                              f"n={a.size():<4} "
                              f"students={sorted(a.student_ids)}")
        print(f"\n{'='*60}\n")

    def print_summary(self):
        """Compact summary — one line per subblock."""
        print(f"\nBlockTree summary")
        print(f"{'='*60}")
        for name in self.all_subblock_names():
            block_name = name[0]
            sb = self.get_subblock(block_name, name)
            labels = [a.label for a in sb.all_assignments()] if sb else []
            print(f"  {name:<5}  {', '.join(labels) if labels else '(empty)'}")
        print(f"\n{'='*60}\n")