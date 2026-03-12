"""
check_data.py
-------------
Data quality gate — run before any scheduling.

Checks:
    1. All student subjects have at least one qualified teacher
    2. Flags bottleneck subjects (only one teacher available)
    3. Prints enrolment summary across grades
    4. Prints conflict matrices for timetable and exam scheduling
"""

from school_tree       import load_from_xlsx
from teacher_tree      import load_teachers_from_xlsx
from conflict_analyser import TimetableConflictAnalyser, ExamConflictAnalyser, MANUAL_SUBJECTS

STUDENT_FILE = "data/students.xlsx"
TEACHER_FILE = "data/teachers.xlsx"

def main():
    # ------------------------------------------------------------------
    # Load
    # ------------------------------------------------------------------
    print(f"Loading students : {STUDENT_FILE}")
    student_tree = load_from_xlsx(STUDENT_FILE)

    print(f"Loading teachers : {TEACHER_FILE}")
    teacher_tree = load_teachers_from_xlsx(TEACHER_FILE)

    # ------------------------------------------------------------------
    # 1. Uncovered subjects
    # ------------------------------------------------------------------
    schedulable = student_tree.all_subjects(exclude=MANUAL_SUBJECTS)
    uncovered   = teacher_tree.uncovered_subjects(list(schedulable))

    print(f"\n{'='*55}")
    print(f"Data quality check")
    print(f"{'='*55}")

    if uncovered:
        print(f"\n  *** Subjects with NO qualified teacher ({len(uncovered)}):")
        for s in uncovered:
            print(f"      {s}")
    else:
        print(f"\n  All schedulable subjects have at least one teacher. OK")

    # ------------------------------------------------------------------
    # 2. Bottleneck subjects
    # ------------------------------------------------------------------
    bottlenecks = teacher_tree.bottleneck_subjects()
    if bottlenecks:
        print(f"\n  *** Bottleneck subjects — only one teacher ({len(bottlenecks)}):")
        for pool in bottlenecks:
            print(f"      {pool.subject_code:<8} -> {pool.teacher_codes()[0]}")
    else:
        print(f"\n  No bottleneck subjects. OK")

    # ------------------------------------------------------------------
    # 3. Enrolment summary
    # ------------------------------------------------------------------
    print(f"\n  Enrolment summary:")
    summary = student_tree.subject_enrolment_summary()
    for grade in sorted(summary.keys()):
        subjects = dict(sorted(summary[grade].items()))
        print(f"    {grade}: {subjects}")

    # ------------------------------------------------------------------
    # 4. Conflict matrices
    # ------------------------------------------------------------------
    ta = TimetableConflictAnalyser(student_tree)
    ta.print_all()
    ta.print_scheduling_order()

    ea = ExamConflictAnalyser(student_tree, merge_grades=True)
    ea.print_all()


if __name__ == "__main__":
    main()