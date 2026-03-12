"""
teacher_tree.py
---------------
OOP model for teacher subject qualifications.

Tree structure:
    TeacherTree (root)
    ├── teachers:      { teacher_code -> Teacher }
    └── subject_pools: { subject_code -> SubjectPool }

Uses xlsx_loader.py for file ingestion.
"""

from xlsx_loader import load_rows, collect_subjects


# ---------------------------------------------------------------------------
# Core classes
# ---------------------------------------------------------------------------

class Teacher:
    def __init__(self, teacher_id: str, code: str, subjects: list[str]):
        self.teacher_id = teacher_id
        self.code       = code
        self.subjects   = subjects

    def can_teach(self, subject: str) -> bool:
        return subject in self.subjects

    def __repr__(self):
        return f"Teacher({self.code}, subjects={self.subjects})"


class SubjectPool:
    """All teachers qualified to teach a given subject."""

    def __init__(self, subject_code: str):
        self.subject_code = subject_code
        self.teachers: list[Teacher] = []

    def add_teacher(self, teacher: Teacher):
        self.teachers.append(teacher)

    def size(self) -> int:
        return len(self.teachers)

    def is_bottleneck(self) -> bool:
        return self.size() == 1

    def teacher_codes(self) -> list[str]:
        return [t.code for t in self.teachers]

    def __repr__(self):
        flag = " *** BOTTLENECK" if self.is_bottleneck() else ""
        return (f"SubjectPool({self.subject_code}, "
                f"n={self.size()}, teachers={self.teacher_codes()}{flag})")


class TeacherTree:
    """Root of the teacher structure — two complementary views."""

    def __init__(self):
        self.teachers:      dict[str, Teacher]     = {}
        self.subject_pools: dict[str, SubjectPool] = {}

    def add_teacher(self, teacher: Teacher):
        self.teachers[teacher.code] = teacher
        for subject in teacher.subjects:
            if subject not in self.subject_pools:
                self.subject_pools[subject] = SubjectPool(subject)
            self.subject_pools[subject].add_teacher(teacher)

    def get_teacher(self, code: str) -> Teacher | None:
        return self.teachers.get(code)

    def get_subject_pool(self, subject: str) -> SubjectPool | None:
        return self.subject_pools.get(subject)

    def all_subjects(self) -> list[str]:
        return sorted(self.subject_pools.keys())

    def all_teacher_codes(self) -> list[str]:
        return sorted(self.teachers.keys())

    def bottleneck_subjects(self) -> list[SubjectPool]:
        return [p for p in self.subject_pools.values() if p.is_bottleneck()]

    def uncovered_subjects(self, school_subjects: list[str]) -> list[str]:
        """Subjects students are enrolled in that have no qualified teacher."""
        return sorted(s for s in school_subjects if s not in self.subject_pools)

    def as_teacher_sets(self) -> dict[str, set[str]]:
        """
        Returns { subject: set_of_teacher_codes }.
        Used by ConflictMatrix to find teacher conflicts —
        two subjects that share no teachers can never run in parallel
        unless a new teacher is hired.
        """
        return {
            subj: set(pool.teacher_codes())
            for subj, pool in self.subject_pools.items()
        }

    def print_tree(self):
        print(f"\nTeacherTree")
        print(f"{'='*55}")

        print(f"\n  By teacher ({len(self.teachers)} total)")
        print(f"  {'-'*45}")
        for code in self.all_teacher_codes():
            t = self.teachers[code]
            print(f"    {t.code:<8} id={t.teacher_id:<4} subjects={t.subjects}")

        print(f"\n  By subject ({len(self.subject_pools)} subjects)")
        print(f"  {'-'*45}")
        for subject in self.all_subjects():
            pool = self.subject_pools[subject]
            flag = "  *** BOTTLENECK" if pool.is_bottleneck() else ""
            print(f"    {subject:<8} n={pool.size():<3} "
                  f"{pool.teacher_codes()}{flag}")

        bottlenecks = self.bottleneck_subjects()
        if bottlenecks:
            print(f"\n  Bottleneck subjects ({len(bottlenecks)}):")
            for pool in bottlenecks:
                print(f"    {pool.subject_code:<8} -> {pool.teacher_codes()[0]}")

        print(f"\n{'='*55}\n")


# ---------------------------------------------------------------------------
# xlsx ingestion — uses generic loader
# ---------------------------------------------------------------------------

def _parse_teacher_row(row: list) -> Teacher | None:
    """
    Row parser for teacher xlsx.
    Col 0: teacher ID, Col 1: teacher code, Col 2+: subject codes
    """
    teacher_id = str(int(row[0]))
    code       = row[1].strip() if row[1] else None

    if not code:
        return None

    subjects = collect_subjects(row, start_col=2)
    return Teacher(teacher_id=teacher_id, code=code, subjects=subjects)


def load_teachers_from_xlsx(filepath: str,
                             sheet_name: str | None = None) -> TeacherTree:
    teachers = load_rows(
        filepath,
        row_parser=_parse_teacher_row,
        sheet_name=sheet_name
    )
    tree = TeacherTree()
    for teacher in teachers:
        tree.add_teacher(teacher)
    return tree


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    filepath = sys.argv[1] if len(sys.argv) > 1 else "data/teachers.xlsx"
    print(f"Loading from: {filepath}")
    tree = load_teachers_from_xlsx(filepath)
    tree.print_tree()