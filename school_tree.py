"""
school_tree.py
--------------
OOP model for student subject enrolments.

Tree structure:
    SchoolTree (root)
    └── GradeNode  (one per grade, e.g. "Gr_8")
        └── SubjectGroup  (one per subject within that grade)
            └── [Student, Student, ...]

Uses xlsx_loader.py for file ingestion.
"""

from xlsx_loader import load_rows, collect_subjects


# ---------------------------------------------------------------------------
# Core classes
# ---------------------------------------------------------------------------

class Student:
    def __init__(self, student_id: str, grade: str, subjects: list[str]):
        self.student_id = student_id
        self.grade      = grade
        self.subjects   = subjects

    def __repr__(self):
        return f"Student({self.student_id}, {self.grade}, {self.subjects})"


class SubjectGroup:
    """All students in a given grade who take a given subject."""

    def __init__(self, subject_code: str, grade: str):
        self.subject_code = subject_code
        self.grade        = grade
        self.students: list[Student] = []

    def add_student(self, student: Student):
        self.students.append(student)

    def member_ids(self) -> set[str]:
        return {s.student_id for s in self.students}

    def size(self) -> int:
        return len(self.students)

    def __repr__(self):
        return (f"SubjectGroup({self.grade}/{self.subject_code}, "
                f"n={self.size()})")


class GradeNode:
    """One grade — holds all SubjectGroups for that grade."""

    def __init__(self, grade: str):
        self.grade = grade
        self.subject_groups: dict[str, SubjectGroup] = {}

    def add_student(self, student: Student):
        for subject in student.subjects:
            if subject not in self.subject_groups:
                self.subject_groups[subject] = SubjectGroup(subject, self.grade)
            self.subject_groups[subject].add_student(student)

    def subjects(self) -> list[str]:
        return sorted(self.subject_groups.keys())

    def get_group(self, subject: str) -> SubjectGroup | None:
        return self.subject_groups.get(subject)

    def as_member_sets(self, exclude: set[str] = None) -> dict[str, set[str]]:
        """
        Returns { subject: set_of_student_ids } for use by ConflictMatrix.
        Excludes manually scheduled subjects if provided.
        """
        exclude = exclude or set()
        return {
            subj: group.member_ids()
            for subj, group in self.subject_groups.items()
            if subj not in exclude
        }

    def __repr__(self):
        return f"GradeNode({self.grade}, subjects={self.subjects()})"


class SchoolTree:
    """Root of the student tree."""

    def __init__(self):
        self.grades: dict[str, GradeNode] = {}

    def add_student(self, student: Student):
        if student.grade not in self.grades:
            self.grades[student.grade] = GradeNode(student.grade)
        self.grades[student.grade].add_student(student)

    def get_grade(self, grade: str) -> GradeNode | None:
        return self.grades.get(grade)

    def get_subject_group(self, grade: str, subject: str) -> SubjectGroup | None:
        node = self.get_grade(grade)
        return node.get_group(subject) if node else None

    def all_grades(self) -> list[str]:
        return sorted(self.grades.keys())

    def all_subjects(self, exclude: set[str] = None) -> set[str]:
        """All unique subject codes across all grades."""
        exclude = exclude or set()
        subjects = set()
        for node in self.grades.values():
            subjects.update(
                s for s in node.subjects() if s not in exclude
            )
        return subjects

    def subject_enrolment_summary(self) -> dict[str, dict[str, int]]:
        """{ grade: { subject: student_count } }"""
        return {
            grade: {
                subj: group.size()
                for subj, group in node.subject_groups.items()
            }
            for grade, node in self.grades.items()
        }

    def print_tree(self):
        print(f"\nSchoolTree")
        print(f"{'='*55}")
        for grade in self.all_grades():
            node = self.grades[grade]
            total = len({
                s.student_id
                for sg in node.subject_groups.values()
                for s in sg.students
            })
            print(f"\n  {grade}  ({total} students)")
            for subject in node.subjects():
                group = node.subject_groups[subject]
                ids   = ", ".join(s.student_id for s in group.students)
                print(f"    {subject:<8} n={group.size():<3}  [{ids}]")
        print(f"\n{'='*55}\n")


# ---------------------------------------------------------------------------
# xlsx ingestion — uses generic loader
# ---------------------------------------------------------------------------

def _normalise_grade(raw: str) -> str:
    return f"Gr_{int(raw)}"


def _parse_student_row(row: list) -> Student | None:
    """
    Row parser for student xlsx.
    Col 0: student ID, Col 1: grade, Col 2+: subject codes
    """
    student_id = str(int(row[0]))
    grade      = _normalise_grade(row[1])
    subjects   = collect_subjects(row, start_col=2)

    if not subjects:
        return None

    return Student(student_id=student_id, grade=grade, subjects=subjects)


def load_from_xlsx(filepath: str) -> SchoolTree:
    students = load_rows(filepath, row_parser=_parse_student_row)
    tree = SchoolTree()
    for student in students:
        tree.add_student(student)
    return tree


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    filepath = sys.argv[1] if len(sys.argv) > 1 else "data/students.xlsx"
    print(f"Loading from: {filepath}")
    tree = load_from_xlsx(filepath)
    tree.print_tree()

    print("Enrolment summary:")
    for grade, subjects in tree.subject_enrolment_summary().items():
        print(f"  {grade}: {dict(sorted(subjects.items()))}")