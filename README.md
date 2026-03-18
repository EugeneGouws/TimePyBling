# TimePyBling

A desktop tool for South African school timetable analysis and exam scheduling. Reads a completed student timetable from Excel, checks it for double-bookings, and builds a dated cross-grade exam timetable with clash-free placement.

## Requirements

- Python 3.13
- `pip install pandas openpyxl reportlab`

## Quick start

```
python main.py
```

Place your data files in `data/` before launching:

| File | Description |
|---|---|
| `data/ST12026.xlsx` | Student timetable — one row per student, columns A1–H7 |
| `data/teachers.xlsx` | Teacher codes and subject pools (optional) |

Each timetable cell contains `SUBJECT TEACHERCODE`, e.g. `AF BALAY`.

---

## The four tabs

### Timetable

Browse the loaded timetable as a tree: Block → SubBlock → Class → Students. A live search box filters by student ID or class label anywhere in the tree.

### Verification

Two checks run against the loaded timetable:

- **Student double-bookings** — any student appearing in two classes in the same subblock.
- **Teacher qualification check** — flags teachers assigned to subjects outside their declared subject pool (requires `teachers.xlsx`).

### Exams

Build a dated exam timetable across all grades in one pass.

**Setting up papers**

The exam tree shows every grade and subject found in the timetable. Each subject starts with one paper (P1). You can:

- Select subjects and add a P2 or P3 paper (up to three papers per subject per grade).
- Tag any paper with a constraint code (e.g. a venue code like `C6`). Papers sharing a constraint code cannot be placed in the same slot, even across grades with no shared students.

Subjects excluded by default: ST, LIB, PE, RDI.

**Generating the schedule**

Set the total number of exam days and a start date, then click Generate. The scheduler places all papers — all grades, all papers — in a single cross-grade pass.

**Viewing and exporting**

A popout window shows the full cross-grade schedule. Filter by grade or view all grades together. Export to PDF or plain text.

### Export

Write the optimised timetable back to ST1.xlsx. (Simulated annealing optimiser pending.)

---

## Scheduling algorithm

All papers across all grades are scheduled simultaneously. Grade-by-grade scheduling is not used because constraint codes create hard clashes between grades even when no students overlap.

**Step 1 — Priority placement (MA and PH)**

Maths and Physical Science papers are placed first, from Gr12 down. For each paper the algorithm picks the valid slot that maximises the gap to other papers of the same subject within the same grade (e.g. MA P1 and MA P2 for Gr12 are spread as far apart as possible).

**Step 2 — DSatur for remaining papers**

Remaining papers are sorted by student count (largest first), then placed using the DSatur graph-colouring heuristic. Papers with more neighbours in the clash graph are assigned first; ties are broken by student count then grade.

**Step 3 — Spacing pass**

After initial placement the algorithm checks whether any same-subject, same-grade papers landed in the same calendar week. Non-priority papers are swapped to a better slot if one exists with no new clashes. Remaining same-week pairs are flagged as warnings in the output.

**Hell week diagnostic**

Any grade where a student has four or more exams in one week is flagged in the output.

**Slots**

Two sessions per day (AM / PM). Weekends are skipped. Slot 0 = Day 1 AM, Slot 1 = Day 1 PM, Slot 2 = Day 2 AM, and so on.

---

## Project structure

```
TimePyBling/
├── main.py
├── data/
├── core/
│   ├── timetable_tree.py      # Parse ST1.xlsx → TimetableTree
│   ├── block_tree.py          # Mutable BlockTree (SA optimiser, pending)
│   ├── conflict_matrix.py     # Shared-student conflict utility
│   └── timetable_converter.py # TimetableTree → BlockTree
└── reader/
│   ├── exam_tree.py           # Reorganise TimetableTree by grade + subject
│   ├── exam_paper.py          # ExamPaper model and ExamPaperRegistry
│   ├── exam_clash.py          # Clash graph — student clashes + constraint clashes
│   ├── exam_scheduler.py      # Cross-grade priority scheduler → ScheduleResult
│   └── verify_timetable.py    # Student double-booking detection
└── ui/
    └── ui.py                  # Tkinter interface (4 tabs)
```

### Key data models

| Class | Module | Role |
|---|---|---|
| `TimetableTree` | `core/timetable_tree.py` | Block → SubBlock → ClassList → StudentList |
| `ExamTree` | `reader/exam_tree.py` | GradeNode → ExamSubject → ClassList |
| `ExamPaper` | `reader/exam_paper.py` | One schedulable exam event (subject + paper number + grade) |
| `ExamPaperRegistry` | `reader/exam_paper.py` | All papers; add/remove P2/P3; constraint codes |
| `ScheduleResult` | `reader/exam_scheduler.py` | Dated slot assignments + warnings |

### Label formats

**Class labels** follow `SUBJECT_TEACHERCODE_GRADE` — e.g. `AF_BALAY_08`, `MA_ALLEN_10`. Grade is zero-padded. `OD` merges to `DR` at parse time.

**Paper labels** follow `SUBJECT_P{n}_GRADE` — e.g. `MA_P1_Gr12`, `EN_P2_Gr10`.

### teachers.xlsx columns

| Column | Description |
|---|---|
| `Teacher Code` | Code used in ST1 cells, e.g. `BALAY` |
| `sua` | Primary subject code |
| `sub` | Secondary subject code (optional) |
| `suc` | Tertiary subject code (optional) |

---

## Backlog

- Inline paper/constraint controls in the exam tree (currently in bottom panel)
- Simulated annealing timetable optimiser
- Per-student printable exam timetable
