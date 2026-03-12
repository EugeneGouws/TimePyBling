"""
conflict_analyser.py
--------------------
Context-aware wrappers around the generic ConflictMatrix.

Two analysers share identical logic — only the data preparation differs:

    TimetableConflictAnalyser
        Input  : SchoolTree
        Scope  : Per grade
        Groups : subject enrolments within one grade
        Use    : Block / slot assignment

    ExamConflictAnalyser
        Input  : SchoolTree
        Scope  : Per grade (or cross-grade for senior exams)
        Groups : subject enrolments, potentially spanning multiple grades
        Use    : Exam timetable construction

Both produce ConflictMatrix objects which expose:
    - conflict_pairs()   subjects that cannot share a slot
    - free_pairs()       subjects that can share a slot
    - ordering()         most-constrained-first scheduling sequence
    - degrees()          conflict count per subject
"""

from conflict_matrix import ConflictMatrix
from school_tree    import SchoolTree


# ---------------------------------------------------------------------------
# Subjects excluded from automated scheduling
# ---------------------------------------------------------------------------

MANUAL_SUBJECTS = {"LIB", "Study", "RDI"}


# ---------------------------------------------------------------------------
# Timetable conflict analyser — per grade
# ---------------------------------------------------------------------------

class TimetableConflictAnalyser:
    """
    Builds one ConflictMatrix per grade from student subject enrolments.
    Used to determine which subjects can share a timetable slot.
    """

    def __init__(self, school_tree: SchoolTree,
                 exclude: set[str] = None):
        self.school_tree = school_tree
        self.exclude     = exclude or MANUAL_SUBJECTS
        self.matrices: dict[str, ConflictMatrix] = {}
        self._build()

    def _build(self):
        for grade in self.school_tree.all_grades():
            node   = self.school_tree.grades[grade]
            groups = node.as_member_sets(exclude=self.exclude)
            self.matrices[grade] = ConflictMatrix(
                label=grade,
                groups=groups
            )

    def get_matrix(self, grade: str) -> ConflictMatrix | None:
        return self.matrices.get(grade)

    def scheduling_order(self) -> list[tuple[str, str, int]]:
        """
        Global scheduling order across all grades.
        Returns (grade, subject, conflict_degree) sorted most-constrained first.
        Feeds directly into the BlockAssigner search.
        """
        order = []
        for grade, matrix in self.matrices.items():
            for subject in matrix.ordering():
                degree = matrix.degrees()[subject]
                order.append((grade, subject, degree))
        return sorted(order, key=lambda x: -x[2])

    def print_all(self):
        print(f"\nTimetableConflictAnalyser")
        print(f"{'='*55}")
        for grade in sorted(self.matrices.keys()):
            self.matrices[grade].print_matrix()
        print(f"\n{'='*55}\n")

    def print_scheduling_order(self):
        print(f"\nGlobal scheduling order (most constrained first):")
        print(f"  {'Grade':<12} {'Subject':<10} {'Degree'}")
        print(f"  {'-'*32}")
        for grade, subject, degree in self.scheduling_order():
            print(f"  {grade:<12} {subject:<10} {degree}")


# ---------------------------------------------------------------------------
# Exam conflict analyser — per grade (or cross-grade)
# ---------------------------------------------------------------------------

class ExamConflictAnalyser:
    """
    Builds one ConflictMatrix per grade for exam scheduling.
    Same logic as TimetableConflictAnalyser — different context.

    For cross-grade exams (e.g. Gr 11 and 12 writing the same paper),
    pass merge_grades=True and the analyser will pool students from
    all grades into one matrix.
    """

    def __init__(self, school_tree: SchoolTree,
                 exclude: set[str] = None,
                 merge_grades: bool = False):
        self.school_tree  = school_tree
        self.exclude      = exclude or MANUAL_SUBJECTS
        self.merge_grades = merge_grades
        self.matrices: dict[str, ConflictMatrix] = {}
        self._build()

    def _build(self):
        if self.merge_grades:
            # Pool all students across all grades into one matrix
            combined: dict[str, set[str]] = {}
            for grade in self.school_tree.all_grades():
                node = self.school_tree.grades[grade]
                for subj, members in node.as_member_sets(self.exclude).items():
                    combined.setdefault(subj, set()).update(members)
            self.matrices["all_grades"] = ConflictMatrix(
                label="all_grades",
                groups=combined
            )
        else:
            # One matrix per grade — same as timetable analyser
            for grade in self.school_tree.all_grades():
                node   = self.school_tree.grades[grade]
                groups = node.as_member_sets(exclude=self.exclude)
                self.matrices[grade] = ConflictMatrix(
                    label=f"exam_{grade}",
                    groups=groups
                )

    def get_matrix(self, grade: str = "all_grades") -> ConflictMatrix | None:
        return self.matrices.get(grade)

    def exam_order(self) -> list[tuple[str, str, int]]:
        """Subjects sorted most-constrained first for exam slot assignment."""
        order = []
        for label, matrix in self.matrices.items():
            for subject in matrix.ordering():
                degree = matrix.degrees()[subject]
                order.append((label, subject, degree))
        return sorted(order, key=lambda x: -x[2])

    def print_all(self):
        print(f"\nExamConflictAnalyser "
              f"({'merged' if self.merge_grades else 'per grade'})")
        print(f"{'='*55}")
        for label in sorted(self.matrices.keys()):
            self.matrices[label].print_matrix()
        print(f"\n{'='*55}\n")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    from school_tree import load_from_xlsx

    filepath = sys.argv[1] if len(sys.argv) > 1 else "data/students.xlsx"

    print(f"Loading students from: {filepath}")
    tree = load_from_xlsx(filepath)

    print("\n--- Timetable conflict analysis ---")
    ta = TimetableConflictAnalyser(tree)
    ta.print_all()
    ta.print_scheduling_order()

    print("\n--- Exam conflict analysis (per grade) ---")
    ea = ExamConflictAnalyser(tree)
    ea.print_all()

    print("\n--- Exam conflict analysis (cross-grade / merged) ---")
    em = ExamConflictAnalyser(tree, merge_grades=True)
    em.print_all()