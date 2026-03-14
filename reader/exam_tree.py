"""
exam_tree.py

Purpose
-------
Take an existing TimetableTree and build a new ExamTree from it.

ExamTree structure
------------------
ExamTree
    -> GradeNode  (Gr08, Gr09, ...)
        -> ExamSubject  (EN_08, AF_08, ...)
            -> ClassList  (EN_COX_08, ...)
                -> StudentList

Example
-------
Gr08
-AF_08
--AF_BALAY_08  (42 students)
---[361, 370, ...]
-EN_08
--EN_COX_08  (22 students)
---[358, 367, ...]

Key design point
----------------
The same class (e.g. AF_BALAY_08) may appear in multiple subblocks
across multiple blocks in the TimetableTree. We merge ALL students
from every appearance into one ClassList here. Set storage means
adding the same student twice is harmless.
"""

from core.timetable_tree import ClassList


# ------------------------------------------------
# EXAM SUBJECT
# ------------------------------------------------
class ExamSubject:
    """
    One subject+grade grouping, e.g. AF_08 or MA_10.
    Contains one ClassList per teacher.
    """

    def __init__(self, label: str):
        self.label = label
        # key = class label (e.g. AF_BALAY_08), value = ClassList
        self.class_lists = {}

    def get_or_create_class_list(self, class_label: str) -> ClassList:
        if class_label not in self.class_lists:
            self.class_lists[class_label] = ClassList(class_label)
        return self.class_lists[class_label]

    def all_students(self) -> set:
        """
        Return the union of all student IDs across every ClassList
        in this subject. Used for clash detection.
        """
        result = set()
        for cl in self.class_lists.values():
            result |= cl.student_list.students
        return result


# ------------------------------------------------
# GRADE NODE
# ------------------------------------------------
class GradeNode:
    """
    Groups all ExamSubjects for one grade.

    Label examples:  Gr08, Gr09, Gr12
    """

    def __init__(self, label: str):
        self.label = label
        # key = exam subject label (e.g. AF_08), value = ExamSubject
        self.exam_subjects = {}

    def get_or_create_exam_subject(self, exam_subject_label: str) -> ExamSubject:
        if exam_subject_label not in self.exam_subjects:
            self.exam_subjects[exam_subject_label] = ExamSubject(exam_subject_label)
        return self.exam_subjects[exam_subject_label]


# ------------------------------------------------
# EXAM TREE
# ------------------------------------------------
class ExamTree:
    """
    Top-level exam tree.

    Structure:
        ExamTree
            -> GradeNode
                -> ExamSubject
                    -> ClassList
                        -> StudentList
    """

    def __init__(self):
        # key = grade label (e.g. 'Gr08'), value = GradeNode
        self.grades = {}

    def get_or_create_grade(self, grade_label: str) -> GradeNode:
        if grade_label not in self.grades:
            self.grades[grade_label] = GradeNode(grade_label)
        return self.grades[grade_label]

    def add_student_to_class(self,
                             grade_label: str,
                             exam_subject_label: str,
                             class_label: str,
                             student_id: int):
        """
        Insert one student into the correct position in the tree.

        Path:  GradeNode -> ExamSubject -> ClassList -> student_id

        Calling this multiple times for the same student is safe
        because StudentList uses a set internally.
        """
        grade        = self.get_or_create_grade(grade_label)
        exam_subject = grade.get_or_create_exam_subject(exam_subject_label)
        class_list   = exam_subject.get_or_create_class_list(class_label)
        class_list.add_student(student_id)

    def print_tree(self):
        for grade_label in sorted(self.grades.keys()):
            grade = self.grades[grade_label]
            print(grade.label)

            for subject_label in sorted(grade.exam_subjects.keys()):
                subject = grade.exam_subjects[subject_label]
                print(f"-{subject.label}")

                for class_label in sorted(subject.class_lists.keys()):
                    cl = subject.class_lists[class_label]
                    print(f"--{cl.label}  ({len(cl.student_list)} students)")
                    print(f"---{cl.student_list}")

            print()


# ------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------
def build_grade_label(grade_str: str) -> str:
    """
    Convert the grade portion of a class label into a GradeNode label.

    Examples:
        "08"  ->  Gr08
        "12"  ->  Gr12
    """
    return f"Gr{grade_str}"


def build_exam_subject_label(class_label: str) -> str:
    """
    Strip the teacher name from a class label, keeping subject + grade.

    Examples:
        AF_BALAY_08         ->  AF_08
        MA_ALLEN_10         ->  MA_10
        DR_VAN_DEN_BERG_12  ->  DR_12
    """
    parts        = class_label.split("_")
    subject_code = parts[0]
    grade        = parts[-1]
    return f"{subject_code}_{grade}"


def build_exam_tree_from_timetable_tree(timetable_tree) -> ExamTree:
    """
    Traverse the entire TimetableTree and build an ExamTree.

    Every block, every subblock, every ClassList is visited so that
    students are merged correctly across all timetable positions.
    """
    exam_tree = ExamTree()

    for block in timetable_tree.blocks.values():
        for subblock in block.subblocks.values():
            for class_list in subblock.class_lists.values():

                class_label        = class_list.label
                exam_subject_label = build_exam_subject_label(class_label)
                grade_str          = class_label.split("_")[-1]
                grade_label        = build_grade_label(grade_str)

                for student_id in class_list.student_list.students:
                    exam_tree.add_student_to_class(
                        grade_label,
                        exam_subject_label,
                        class_label,
                        student_id
                    )

    return exam_tree


# ------------------------------------------------
# OPTIONAL TEST NOTE
# ------------------------------------------------
if __name__ == "__main__":
    print("This file is meant to be used from main.py")