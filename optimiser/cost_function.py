"""
cost_function.py
----------------
Evaluates the timetable cost function E(T):

  E(T) = 10000·C_s + 5000·C_t + 400·P_g12 + 100·P_tg + 50·P_f + 10·P_stg + 5·P_alloc

Term definitions
----------------
Hard
  C_s    student clashes    — student in 2+ classes in the same subblock
  C_t    teacher clashes    — teacher in 2+ real-teaching classes in the same
                               subblock (supervision subjects LIB/Study excluded)

Preference  (stub — require optional columns in teachers.xlsx)
  P_g12  Gr 12 teacher pref  — Gr 12 class taught by non-preferred teacher
  P_tg   teacher grade pref  — teacher assigned outside their preferred grade set
  P_f    teacher free-day    — teacher has a lesson on their preferred free day

Soft structural
  P_stg  sparse staggering   — a (subject, block, grade) group with few lessons
                               appears on consecutive cycle days
  P_alloc allocation         — stub for future Gr 9/10/11 allocation rules

Dropped terms
-------------
  C_spd  subject-once-per-day  — dropped; flex-block structure makes this
                                  indistinguishable from legitimate doubles
  C_lb   linked-block          — removed; MA/LO alternation logic was not
                                  reliable enough to use as a constraint

Timetable coordinate system
----------------------------
  Subblock name: letter + number  e.g. "A3", "G5"
  Block letter = which lesson slot within the day  (A–H)
  Subblock num = cycle day the slot falls on (1–8)

Usage
-----
    from optimiser.cost_function import evaluate, CostConfig

    result = evaluate(tree)
    print(result)
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

from core.block_tree import BlockTree


# ─────────────────────────────────────────────────────────────────────────────
# Label parsing helpers
# ─────────────────────────────────────────────────────────────────────────────

def _subject(label: str) -> str:
    return label.split("_")[0]

def _grade(label: str) -> str:
    return label.split("_")[-1]

def _teacher(label: str) -> str:
    parts = label.split("_")
    return "_".join(parts[1:-1])

def _day(subblock_name: str) -> int:
    return int(subblock_name[1:])

def _block(subblock_name: str) -> str:
    return subblock_name[0]


# ─────────────────────────────────────────────────────────────────────────────
# Configuration dataclasses
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class CostWeights:
    """Weights for each term in E(T). Adjust to tune optimiser priorities."""
    student_clash: int = 10_000   # C_s
    teacher_clash: int =  5_000   # C_t
    gr12_pref:     int =    400   # P_g12
    teacher_grade: int =    100   # P_tg
    teacher_free:  int =     50   # P_f
    stagger:       int =     10   # P_stg
    alloc:         int =      5   # P_alloc


@dataclass
class TeacherPreferences:
    """
    Optional teacher preference data. All fields default to empty,
    which disables the corresponding preference terms.

    gr12_teachers
        Teacher codes preferred for Gr 12. P_g12 fires for Gr 12 classes
        taught by anyone NOT in this set.

    teacher_grades
        teacher_code -> frozenset of preferred grade strings e.g. {"10","11","12"}.
        P_tg fires when a teacher is assigned outside their preferred grades.

    teacher_free_days
        teacher_code -> frozenset of preferred free day numbers (1-8).
        P_f fires when a teacher has a lesson on one of their free days.
    """
    gr12_teachers:     set[str]             = field(default_factory=set)
    teacher_grades:    dict[str, frozenset] = field(default_factory=dict)
    teacher_free_days: dict[str, frozenset] = field(default_factory=dict)


@dataclass
class CostConfig:
    """
    Full configuration for the cost evaluator.
    evaluate(tree) works with no arguments — all defaults are sensible.
    """
    weights:       CostWeights        = field(default_factory=CostWeights)
    teacher_prefs: TeacherPreferences = field(default_factory=TeacherPreferences)

    # Subjects excluded from teacher-clash checking.
    # LIB and Study are supervision — one teacher may cover multiple groups.
    teacher_clash_exclude: set[str] = field(
        default_factory=lambda: {"LIB", "Study"}
    )

    # Lesson count at or below which a (subject, block, grade) group is
    # considered sparse and checked for consecutive-day staggering.
    # Set to 0 to disable P_stg entirely.
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
    C_s:     int = 0
    C_t:     int = 0
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
        W = 28

        def row(name, val, stub=False):
            suffix = "  *stub*" if stub else ""
            return f"  {name:<{W}} {val:>6}{suffix}"

        lines = [
            "",
            "=" * 54,
            "  COST FUNCTION BREAKDOWN",
            "=" * 54,
            row("C_s   student clashes",    self.C_s),
            row("C_t   teacher clashes",    self.C_t),
            row("P_g12 Gr 12 teacher pref", self.P_g12,  stub=True),
            row("P_tg  teacher grade pref", self.P_tg,   stub=True),
            row("P_f   teacher free day",   self.P_f,    stub=True),
            row("P_stg sparse staggering",  self.P_stg),
            row("P_alloc allocation",       self.P_alloc, stub=True),
            "  " + "-" * 40,
            row("E(T)  TOTAL",              self.total),
            "=" * 54,
            f"  Feasible (no hard clashes): {self.is_feasible()}",
            "",
        ]
        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# Index builder
# ─────────────────────────────────────────────────────────────────────────────

def _build_indices(tree: BlockTree) -> dict:
    """Single pass through the tree building all lookup structures."""
    subblock_students: dict = defaultdict(lambda: defaultdict(list))
    subblock_teachers: dict = defaultdict(lambda: defaultdict(list))
    teacher_day_sbs:   dict = defaultdict(lambda: defaultdict(list))
    sbg_days:          dict = defaultdict(list)

    for block in tree.blocks.values():
        for sb_name, subblock in block.subblocks.items():
            day = _day(sb_name)
            blk = _block(sb_name)

            for assignment in subblock.all_assignments():
                label   = assignment.label
                subj    = _subject(label)
                grade   = _grade(label)
                teacher = _teacher(label)

                for sid in assignment.student_ids:
                    subblock_students[sb_name][sid].append(label)

                subblock_teachers[sb_name][teacher].append(label)
                teacher_day_sbs[teacher][day].append(sb_name)
                sbg_days[(subj, blk, grade)].append(day)

    for key in sbg_days:
        sbg_days[key].sort()

    return {
        "subblock_students": subblock_students,
        "subblock_teachers": subblock_teachers,
        "teacher_day_sbs":   teacher_day_sbs,
        "sbg_days":          sbg_days,
    }

# ─────────────────────────────────────────────────────────────────────────────
# Term evaluators
# ─────────────────────────────────────────────────────────────────────────────

def _eval_student_clashes(idx: dict) -> int:
    total = 0
    for sid_map in idx["subblock_students"].values():
        for classes in sid_map.values():
            if len(classes) > 1:
                total += len(classes) - 1
    return total


def _eval_teacher_clashes(idx: dict, config: CostConfig) -> int:
    total = 0
    for teacher_map in idx["subblock_teachers"].values():
        for teacher, classes in teacher_map.items():
            real = [c for c in classes
                    if _subject(c) not in config.teacher_clash_exclude]
            if len(real) > 1:
                total += len(real) - 1
    return total


def _eval_gr12_teacher_pref(idx: dict, config: CostConfig) -> int:
    """STUB: returns 0 until teacher_prefs.gr12_teachers is populated."""
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
    """STUB: returns 0 until teacher_prefs.teacher_grades is populated."""
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
    """STUB: returns 0 until teacher_prefs.teacher_free_days is populated."""
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
    """P_stg — sparse groups on consecutive cycle days. Violation = days d and d+1."""
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
    """STUB: returns 0 until Gr 9/10/11 allocation rules are defined."""
    return 0


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def evaluate(
    tree:   BlockTree,
    config: Optional[CostConfig] = None,
) -> CostBreakdown:
    """
    Compute E(T) for the given TimetableTree.

    Parameters
    ----------
    tree   : populated TimetableTree
    config : CostConfig — all defaults used if None

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
        P_g12   = _eval_gr12_teacher_pref(idx, config),
        P_tg    = _eval_teacher_grade_pref(idx, config),
        P_f     = _eval_teacher_free_pref(idx, config),
        P_stg   = _eval_stagger(idx, config),
        P_alloc = _eval_alloc(idx, config),
    )

    result.total = (
        w.student_clash * result.C_s   +
        w.teacher_clash * result.C_t   +
        w.gr12_pref     * result.P_g12 +
        w.teacher_grade * result.P_tg  +
        w.teacher_free  * result.P_f   +
        w.stagger       * result.P_stg +
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
        gr12        (Y/N)  — preferred for Gr 12
        pref_grades (str)  — comma-separated grades e.g. "10,11,12"
        free_days   (str)  — comma-separated day numbers e.g. "5,6"
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
                int(d.strip())
                for d in str(fd).split(",")
                if d.strip().isdigit()
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

    st_file = Path("data/ST1.xlsx")
    if not st_file.exists():
        print("data/ST1.xlsx not found.")
        raise SystemExit(1)

    print(f"Loading: {st_file}")
    tree   = build_timetable_tree_from_file(st_file)
    result = evaluate(tree)
    print(result)