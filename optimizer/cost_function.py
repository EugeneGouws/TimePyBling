"""
cost_function.py
----------------
Evaluates the timetable cost function E(T):

  E(T) = 10000·C_s  + 5000·C_t  + 1000·C_spd
       +   400·C_lb +  400·P_g12 +  100·P_tg
       +    50·P_f  +   10·P_stg +    5·P_alloc

(C_md has been removed — the G-block MA lessons are treated as independent
 scheduled classes and their staggering is handled by P_stg.)

Term definitions
----------------
Hard / near-hard
  C_s    student clashes       — student in 2+ classes in the same subblock
  C_t    teacher clashes       — teacher in 2+ real-teaching classes in the same
                                  subblock (supervision subjects LIB/Study excluded)
  C_spd  subject-once-per-day  — a student has the same subject in 2+ slots on
                                  the same cycle day.  Subjects in linked_subjects
                                  (e.g. MA, which legitimately doubles) are fully
                                  excluded from this check.

Structural
  C_lb   linked-block          — student has BOTH subjects in a linked pair
                                  (e.g. MA and LO) on the same cycle day.
                                  On MA-double days the LO slot is displaced by MA,
                                  so co-occurrence is a genuine structural violation.

Preference  (require TeacherPreferences data)
  P_g12  Gr 12 teacher pref    — Gr 12 class taught by non-preferred teacher
  P_tg   teacher grade pref    — teacher assigned a grade outside their preference
  P_f    teacher free-day pref — teacher has a lesson on their preferred free day

Soft structural
  P_stg  sparse staggering     — a (subject, block, grade) group with few lessons
                                  in the cycle appears on consecutive subblock
                                  numbers.

                                  Grouping is by (subject, block_letter, grade)
                                  so G-block MA (3/cycle) and G-block LO (4/cycle)
                                  are checked independently.

                                  Violation = two appearances at subblock numbers
                                  d and d+1 (consecutive days in the same slot).

  P_alloc allocation penalty   — stub for future Gr 9/10/11 allocation rules

Timetable coordinate system
----------------------------
  Subblock name:  letter + number   e.g. "A3", "G5"
  Block letter  = which lesson slot within the day  (A, B, C … H)
  Subblock num  = cycle day this slot falls on (1 – 8)

  A student's full-cycle schedule for a regular subject fills one block
  column (e.g. EN in block B → B1, B2 … B7, one free day).
  The flexible G-block varies by day — different subjects placed there
  on different days of the cycle (e.g. MA on days 2,5,7; LO on the rest).

Usage
-----
    from cost_function import evaluate, CostConfig

    result = evaluate(tree)            # all defaults
    print(result)

    config = CostConfig()
    config.sparse_threshold = 4
    result = evaluate(tree, config)
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

from core.timetable_tree import TimetableTree


# ─────────────────────────────────────────────────────────────────────────────
# Label parsing helpers
# ─────────────────────────────────────────────────────────────────────────────

def _subject(label: str) -> str:
    """'MA_ALLEN_10'  →  'MA'"""
    return label.split("_")[0]


def _grade(label: str) -> str:
    """'MA_ALLEN_10'  →  '10'"""
    return label.split("_")[-1]


def _teacher(label: str) -> str:
    """'DR_VAN_DEN_BERG_12'  →  'VAN_DEN_BERG'"""
    parts = label.split("_")
    return "_".join(parts[1:-1])


def _day(subblock_name: str) -> int:
    """Cycle day from subblock name.  'A3' → 3,  'G7' → 7."""
    return int(subblock_name[1:])


def _block(subblock_name: str) -> str:
    """Block letter from subblock name.  'G5' → 'G'."""
    return subblock_name[0]


# ─────────────────────────────────────────────────────────────────────────────
# Configuration dataclasses
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class CostWeights:
    """
    Weights for each term in E(T).
    Adjust here to tune optimiser priorities without touching the logic.
    """
    student_clash: int = 10_000   # C_s
    teacher_clash: int =  5_000   # C_t
    subj_per_day:  int =      0   # C_spd  (dropped — C_s + C_lb cover the real cases)
    linked_block:  int =    400   # C_lb
    gr12_pref:     int =    400   # P_g12
    teacher_grade: int =    100   # P_tg
    teacher_free:  int =     50   # P_f
    stagger:       int =     10   # P_stg
    alloc:         int =      5   # P_alloc


@dataclass
class TeacherPreferences:
    """
    Teacher preference data for the soft preference terms.
    All fields default to empty (terms disabled) until populated.

    gr12_teachers
        Set of teacher codes preferred for Gr 12 classes.
        P_g12 fires for any Gr 12 class taught by a teacher NOT in this set.

    teacher_grades
        teacher_code → frozenset of preferred grade strings (e.g. {"10","11","12"}).
        P_tg fires when a teacher is assigned a grade outside their preferred set.

    teacher_free_days
        teacher_code → frozenset of preferred free day numbers (1–8).
        P_f fires when a teacher has any class on one of their preferred free days.

    Example
    -------
        prefs = TeacherPreferences(
            gr12_teachers={"ALLEN", "SMITH"},
            teacher_grades={"ALLEN": frozenset({"11", "12"})},
            teacher_free_days={"ALLEN": frozenset({5})},
        )
    """
    gr12_teachers:     set[str]             = field(default_factory=set)
    teacher_grades:    dict[str, frozenset] = field(default_factory=dict)
    teacher_free_days: dict[str, frozenset] = field(default_factory=dict)


@dataclass
class CostConfig:
    """
    Full configuration for the cost evaluator.
    All fields have sensible defaults — evaluate(tree) works with no arguments.
    """
    weights:       CostWeights        = field(default_factory=CostWeights)
    teacher_prefs: TeacherPreferences = field(default_factory=TeacherPreferences)

    teacher_clash_exclude: set[str] = field(
        default_factory=lambda: {"LIB", "ST", "RDI"}
    )

    # Subjects that legitimately run as doubles across the cycle (MA has 10
    # lessons: once per day in its home block + 3 extras in the G-block).
    # These subjects are excluded entirely from C_spd checking.
    linked_subjects: set[str] = field(default_factory=lambda: {"MA"})

    # Subject pairs whose flexible-block appearances alternate.
    # MA displaces LO on 3 days in the G-block.
    # C_lb fires when a student has both subjects on the same cycle day.
    linked_pairs: list[tuple[str, str]] = field(
        default_factory=lambda: [("MA", "LO")]
    )

    # Subjects excluded from teacher-clash checking.
    # LIB and Study are supervision/admin — one teacher may cover multiple groups.
    teacher_clash_exclude: set[str] = field(
        default_factory=lambda: {"LIB", "Study"}
    )

    # Lesson count at or below which a (subject, block, grade) group is
    # considered sparse and should be staggered across the cycle.
    # E.g. a group with 3 G-block lessons should skip at least one day between
    # appearances.  Set to 0 to disable P_stg entirely.
    sparse_threshold: int = 6

    # Grades evaluated for P_alloc (stub — no behaviour yet).
    alloc_grades: set[str] = field(
        default_factory=lambda: {"09", "10", "11"}
    )


# ─────────────────────────────────────────────────────────────────────────────
# Result dataclass
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class CostBreakdown:
    """
    Per-term violation counts and weighted total E(T).
    All counts are non-negative integers.
    """
    C_s:     int = 0
    C_t:     int = 0
    C_spd:   int = 0
    C_lb:    int = 0
    P_g12:   int = 0
    P_tg:    int = 0
    P_f:     int = 0
    P_stg:   int = 0
    P_alloc: int = 0
    total:   int = 0

    def is_feasible(self) -> bool:
        """True when all hard constraints are satisfied (C_s = C_t = 0)."""
        return self.C_s == 0 and self.C_t == 0

    def __str__(self) -> str:
        W         = 28
        stub_note = " *stub*"

        def row(name, val, stub=False):
            return f"  {name:<{W}} {val:>6}{stub_note if stub else ''}"

        lines = [
            "",
            "=" * 54,
            "  COST FUNCTION BREAKDOWN",
            "=" * 54,
            row("C_s   student clashes",    self.C_s),
            row("C_t   teacher clashes",    self.C_t),
            row("C_lb  linked block",       self.C_lb),
            row("P_g12 Gr 12 teacher pref", self.P_g12,   stub=True),
            row("P_tg  teacher grade pref", self.P_tg,    stub=True),
            row("P_f   teacher free day",   self.P_f,     stub=True),
            row("P_stg sparse staggering",  self.P_stg),
            row("P_alloc allocation",       self.P_alloc, stub=True),
            "  " + "─" * 40,
            row("E(T)  TOTAL",              self.total),
            "=" * 54,
            f"  Feasible (no hard clashes): {self.is_feasible()}",
            "",
        ]
        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# Index builder  (single pass through the tree)
# ─────────────────────────────────────────────────────────────────────────────

def _build_indices(tree: TimetableTree) -> dict:
    """
    Walk the tree once and build every lookup structure the term evaluators need.

    Keys
    ----
    subblock_students  : sb_name → {student_id: [label, ...]}
    subblock_teachers  : sb_name → {teacher:    [label, ...]}
    student_day_subj_n : (student_id, day) → {subject: count}
    student_day_subjs  : (student_id, day) → {subject, ...}
    teacher_day_sbs    : teacher → {day: [sb_name, ...]}
    sbg_days           : (subject, block_letter, grade) → sorted [day, ...]
    """
    subblock_students:  dict = defaultdict(lambda: defaultdict(list))
    subblock_teachers:  dict = defaultdict(lambda: defaultdict(list))
    student_day_subj_n: dict = defaultdict(lambda: defaultdict(int))
    student_day_subjs:  dict = defaultdict(set)
    teacher_day_sbs:    dict = defaultdict(lambda: defaultdict(list))
    sbg_days:           dict = defaultdict(list)

    for _block_name, block in tree.blocks.items():
        for sb_name, subblock in block.subblocks.items():
            day = _day(sb_name)
            blk = _block(sb_name)

            for label, cl in subblock.class_lists.items():
                subj    = _subject(label)
                grade   = _grade(label)
                teacher = _teacher(label)

                for sid in cl.student_list.students:
                    subblock_students[sb_name][sid].append(label)
                    student_day_subj_n[(sid, day)][subj] += 1
                    student_day_subjs[(sid, day)].add(subj)

                subblock_teachers[sb_name][teacher].append(label)
                teacher_day_sbs[teacher][day].append(sb_name)
                sbg_days[(subj, blk, grade)].append(day)

    for key in sbg_days:
        sbg_days[key].sort()

    return {
        "subblock_students":  subblock_students,
        "subblock_teachers":  subblock_teachers,
        "student_day_subj_n": student_day_subj_n,
        "student_day_subjs":  student_day_subjs,
        "teacher_day_sbs":    teacher_day_sbs,
        "sbg_days":           sbg_days,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Individual term evaluators
# ─────────────────────────────────────────────────────────────────────────────

def _eval_student_clashes(idx: dict) -> int:
    """
    C_s — student appears in 2+ classes within the same subblock.
    Count = total excess appearances across all subblocks.
    """
    total = 0
    for sid_map in idx["subblock_students"].values():
        for classes in sid_map.values():
            if len(classes) > 1:
                total += len(classes) - 1
    return total


def _eval_teacher_clashes(idx: dict, config: CostConfig) -> int:
    """
    C_t — teacher appears in 2+ real-teaching classes in the same subblock.

    Subjects in teacher_clash_exclude (LIB, Study) are skipped because one
    teacher legitimately supervises multiple concurrent groups there.
    """
    total = 0
    for teacher_map in idx["subblock_teachers"].values():
        for teacher, classes in teacher_map.items():
            real = [c for c in classes
                    if _subject(c) not in config.teacher_clash_exclude]
            if len(real) > 1:
                total += len(real) - 1
    return total


def _eval_subj_per_day(_idx: dict, _config: CostConfig) -> int:
    """
    C_spd — dropped.

    In this school's flex-block structure every subject that uses a flexible
    slot (EN, LO, GE, EMS, HI, …) legitimately appears in two different block
    letters on the same cycle day.  There is no meaningful threshold that
    distinguishes intended doubles from errors without enumerating every
    subject's flex pattern.

    The real hard constraint (student in two slots at the same time) is
    already caught by C_s.  The MA/LO alternation is caught by C_lb.

    Returns 0 always.  Weight is also set to 0 in CostWeights.
    """
    return 0


def _eval_linked_block(idx: dict, config: CostConfig) -> int:
    """
    C_lb — student has BOTH subjects in a linked pair on the same cycle day.

    On MA-double days the G-block MA lesson is supposed to displace LO.
    A student showing both MA and LO on that day means the displacement
    did not happen — a genuine structural violation.

    Count = (student_id, day) pairs where both subjects appear.
    """
    total = 0
    for (sid, day), subjects in idx["student_day_subjs"].items():
        for subj_a, subj_b in config.linked_pairs:
            if subj_a in subjects and subj_b in subjects:
                total += 1
    return total


def _eval_gr12_teacher_pref(idx: dict, config: CostConfig) -> int:
    """
    P_g12 — Gr 12 class taught by a teacher not in the preferred set.
    STUB: returns 0 until teacher_prefs.gr12_teachers is populated.
    """
    if not config.teacher_prefs.gr12_teachers:
        return 0
    total = 0
    for teacher_map in idx["subblock_teachers"].values():
        for teacher, classes in teacher_map.items():
            if teacher in config.teacher_prefs.gr12_teachers:
                continue
            for label in classes:
                if _grade(label) == "12":
                    total += 1
    return total


def _eval_teacher_grade_pref(idx: dict, config: CostConfig) -> int:
    """
    P_tg — teacher assigned a grade outside their preferred grade set.
    STUB: returns 0 until teacher_prefs.teacher_grades is populated.
    """
    if not config.teacher_prefs.teacher_grades:
        return 0
    total = 0
    for teacher_map in idx["subblock_teachers"].values():
        for teacher, classes in teacher_map.items():
            preferred = config.teacher_prefs.teacher_grades.get(teacher)
            if preferred is None:
                continue
            for label in classes:
                if _grade(label) not in preferred:
                    total += 1
    return total


def _eval_teacher_free_pref(idx: dict, config: CostConfig) -> int:
    """
    P_f — teacher has a lesson on one of their preferred free days.
    STUB: returns 0 until teacher_prefs.teacher_free_days is populated.
    """
    if not config.teacher_prefs.teacher_free_days:
        return 0
    total = 0
    for teacher, day_map in idx["teacher_day_sbs"].items():
        preferred_free = config.teacher_prefs.teacher_free_days.get(teacher)
        if not preferred_free:
            continue
        for day in day_map:
            if day in preferred_free:
                total += 1
    return total


def _eval_stagger(idx: dict, config: CostConfig) -> int:
    """
    P_stg — sparse (subject, block, grade) groups on consecutive cycle days.

    A group is sparse when it has <= config.sparse_threshold lessons in the
    cycle.  For each pair of consecutive day numbers (d, d+1) within the
    group, one violation is counted.

    Grouping by (subject, block_letter, grade) means G-block MA (3 lessons)
    and G-block LO (≈4 lessons) are checked independently from A-block MA
    (7 lessons), which is large enough to not be sparse.

    Example
      LO in G1, G2, G4  →  days [1,2,4]  →  1 violation  (1+2 consecutive)
      LO in G1, G3, G5  →  days [1,3,5]  →  0 violations
    """
    if config.sparse_threshold <= 0:
        return 0
    total = 0
    for (_subj, _blk, _grade), days in idx["sbg_days"].items():
        if len(days) > config.sparse_threshold:
            continue
        for i in range(len(days) - 1):
            if days[i + 1] - days[i] == 1:
                total += 1
    return total


def _eval_alloc(_idx: dict, _config: CostConfig) -> int:
    """
    P_alloc — allocation penalties for Gr 9/10/11.
    STUB: returns 0 until rules are defined.
    """
    return 0


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def evaluate(
    tree:   TimetableTree,
    config: Optional[CostConfig] = None,
) -> CostBreakdown:
    """
    Compute E(T) for the given TimetableTree.

    Parameters
    ----------
    tree   : populated TimetableTree (from build_timetable_tree_from_file)
    config : CostConfig — uses all defaults if None

    Returns
    -------
    CostBreakdown with per-term counts and weighted total E(T).
    """
    if config is None:
        config = CostConfig()

    idx = _build_indices(tree)
    w   = config.weights

    result = CostBreakdown(
        C_s     = _eval_student_clashes(idx),
        C_t     = _eval_teacher_clashes(idx, config),
        C_spd   = _eval_subj_per_day(idx, config),
        C_lb    = _eval_linked_block(idx, config),
        P_g12   = _eval_gr12_teacher_pref(idx, config),
        P_tg    = _eval_teacher_grade_pref(idx, config),
        P_f     = _eval_teacher_free_pref(idx, config),
        P_stg   = _eval_stagger(idx, config),
        P_alloc = _eval_alloc(idx, config),
    )

    result.total = (
        w.student_clash * result.C_s     +
        w.teacher_clash * result.C_t     +
        w.subj_per_day  * result.C_spd   +
        w.linked_block  * result.C_lb    +
        w.gr12_pref     * result.P_g12   +
        w.teacher_grade * result.P_tg    +
        w.teacher_free  * result.P_f     +
        w.stagger       * result.P_stg   +
        w.alloc         * result.P_alloc
    )

    return result


# ─────────────────────────────────────────────────────────────────────────────
# Teacher preference loader
# ─────────────────────────────────────────────────────────────────────────────

def load_teacher_prefs_from_xlsx(path: str) -> TeacherPreferences:
    """
    Load teacher preferences from teachers.xlsx.

    Required column : Teacher Code
    Optional columns (add to teachers.xlsx to activate preference terms):
        gr12        (Y/N/bool)  — preferred for Gr 12
        pref_grades (str)       — comma-separated grades, e.g. "10,11,12"
        free_days   (str)       — comma-separated day numbers, e.g. "5,6"

    Missing optional columns leave the corresponding terms stubbed at 0.
    """
    import pandas as pd

    df    = pd.read_excel(path)
    prefs = TeacherPreferences()

    for _, row in df.iterrows():
        code = str(row.get("Teacher Code", "")).strip()
        if not code:
            continue

        gr12_val = row.get("gr12", None)
        if gr12_val is not None:
            if str(gr12_val).strip().upper() in ("Y", "YES", "1", "TRUE"):
                prefs.gr12_teachers.add(code)

        pg = row.get("pref_grades", None)
        if pg is not None and str(pg).strip():
            grades = frozenset(
                g.strip().zfill(2) for g in str(pg).split(",") if g.strip()
            )
            if grades:
                prefs.teacher_grades[code] = grades

        fd = row.get("free_days", None)
        if fd is not None and str(fd).strip():
            days = frozenset(
                int(d.strip()) for d in str(fd).split(",") if d.strip().isdigit()
            )
            if days:
                prefs.teacher_free_days[code] = days

    return prefs


# ─────────────────────────────────────────────────────────────────────────────
# Standalone runner
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    from pathlib import Path
    from core.timetable_tree import build_timetable_tree_from_file

    st_file = None
    for candidate in ["data/ST1.xlsx"]:
        p = Path(candidate)
        if p.exists():
            st_file = p
            break

    if st_file is None:
        print("No ST1.xlsx found in current folder.")
        raise SystemExit(1)

    print(f"Loading: {st_file}")
    tree = build_timetable_tree_from_file(st_file)

    config = CostConfig()
    teacher_file = Path("../data/teachers.xlsx")
    if teacher_file.exists():
        print(f"Loading teacher preferences: {teacher_file}")
        config.teacher_prefs = load_teacher_prefs_from_xlsx(str(teacher_file))

    result = evaluate(tree, config)
    print(result)