# TimePyBling

A desktop tool for South African school timetable analysis and exam scheduling.
Reads a student timetable, checks it for double-bookings, and builds a dated
cross-grade exam timetable with clash-free placement.

## Requirements

- Python 3.13
- `pip install pandas openpyxl reportlab`

> `reportlab` is optional — the schedule can be saved as plain text if not installed.

## Quick start

```
python main.py
```

---

## Data files

TimePyBling does not ship with student data. On first use, load your timetable
from Excel. After configuring papers and dates, save a **state file** so the app
can reload without the original Excel file.

### First-time setup (Excel)

Place the student timetable in `data/` before launching:

| File | Description |
|---|---|
| `data/ST12026.xlsx` | Student timetable — one row per student, columns A1–H7 |

Each timetable cell contains `SUBJECT TEACHERCODE`, e.g. `AF BALAY`.

Auto-load tries these filenames in order:
```
data/ST1 2026.xlsx
data/ST12026.xlsx
data/ST1_2026.xlsx
data/ST1.xlsx
```

### Deployed / subsequent use (JSON state file)

After configuring the exam setup, click **Save State…** and save to
`data/timetable_state.json`. On subsequent launches the app loads this file
automatically — no Excel file required.

The state file stores the full timetable tree, exclusions, papers (P1/P2/P3),
constraint codes, difficulty settings, linked papers, and session date settings.

---

## The three tabs

### Timetable

Browse the loaded timetable as a tree: Block → SubBlock → Class → Students.
Click any class to see its student list. A live search box filters by student
ID or class label anywhere in the tree.

### Verification

Two checks run against the loaded timetable:

- **Student double-bookings** — any student appearing in two classes in the
  same subblock.
- **Teacher qualification check** — flags teachers assigned to subjects outside
  their declared subject pool (requires a teachers file loaded separately).

This tab is shown automatically after any timetable load.

### Exams

Build a dated exam timetable across all grades in one pass.

**Setting up papers**

The exam tab shows a subject × grade matrix. Each cell shows the papers active
for that subject in that grade. Click any cell to open a detail popout where you can:

- Add or remove papers (P1/P2/P3)
- Set difficulty (Red / Yellow / Green) — subject-wide
- Add constraint codes (e.g. venue codes like `C6`)
- Link papers (e.g. Music P1 + P2 placed on the same day)
- Pin a paper to a specific date and session

Subjects excluded by default: `ST`, `LIB`, `PE`, `RDI`.

Use the **[ST+]** button to add Study blocks — these pin a grade to a specific
date/session as a hard slot reservation.

**Difficulty rules**

- Red papers are placed at least 5 days apart (per grade). This guarantees
  zero student stress for students who only sit red papers.
- Yellow and Green papers fill the gaps between red papers.
- AM sessions are always filled before PM sessions are unlocked.
- Linked subjects are auto-promoted to red.

**Linked papers**

Linked papers (e.g. Music P1 and P2) are placed consecutively — P1 in AM,
P2 in PM of the same day. Linked partners are exempt from difficulty clash
rules with each other.

**Generating the schedule**

Set a start date, end date, and AM/PM toggles in the left panel, then click
**Generate**. The scheduler places all papers across all grades in a single pass.

**Teacher optimisation**

After generating, use the **Optimise** button to reduce teacher marking load.
The tolerance slider controls how much student stress may increase:
0% = no change allowed, 100% = optimise for teachers regardless of students.
Click **Generate Schedule** again to reset to the student-optimal result.

**Saving and restoring state**

Use **Save State…** / **Load State…** to save or restore the full configuration
to a JSON file.

---

## Scheduling algorithm

Papers across all grades are placed using a three-phase constructive algorithm.

**Step 0 — Pinned paper reservation**
Pinned ST (study) papers and any manually pinned papers are placed first,
reserving grade/slot pairs that no other paper for that grade may use.

**Phase 1 — Red + Linked (AM only, 5-day spacing)**
Red papers and linked papers (auto-promoted to red) are placed grade by grade
from Gr12 down, most constrained first. Only AM slots are used. A minimum
5-day gap is enforced between same-grade papers, guaranteeing zero overlap
cost for students with only red subjects.
Linked pairs: primary paper placed in AM, partner auto-placed in PM of the
same day.

**Phase 2 — Yellow (fill gaps)**
Yellow papers fill the slots between red papers. AM slots are used first;
PM slots unlock only once all AM slots are exhausted. Each paper is placed
in the slot that minimises incremental per-student overlap cost.

**Phase 3 — Green (fill remaining)**
Same logic as Phase 2 for remaining green papers.

**Post — Hill-climb**
A paper-move local search (20 passes) swaps non-pinned papers to reduce
the per-student overlap cost score.

**Cost function**
Per-student overlap cost: penalises only when a student has 2+ exams in the
same convolution window (window sizes 2, 3, 4 days with weights 3, 2, 1).
Red papers 5+ days apart score zero by construction.

**Slots**
Two sessions per day (AM / PM). Weekends are skipped.
Slot 0 = Day 1 AM, Slot 1 = Day 1 PM, Slot 2 = Day 2 AM, and so on.

---

## Project structure

```
TimePyBling/
├── main.py
├── data/
│   └── timetable_state.json     # generated by Save State — loaded on startup
├── core/
│   ├── timetable_tree.py        # Parse ST1.xlsx → TimetableTree; JSON serialisation
│   ├── conflict_matrix.py       # Shared-student conflict utility
│   ├── cost_function.py         # StudentStressCost, TeacherMarkingCost, TotalCost
│   ├── hill_climb.py            # Slot-swap and paper-move hill-climb optimisers
│   └── cpsat_optimiser.py       # CP-SAT teacher optimiser (future use)
├── reader/
│   ├── exam_tree.py             # Reorganise TimetableTree by grade + subject
│   ├── exam_paper.py            # ExamPaper model and ExamPaperRegistry
│   ├── exam_clash.py            # Clash graph — student + constraint clashes
│   ├── exam_scheduler.py        # Cross-grade priority scheduler → ScheduleResult
│   └── verify_timetable.py      # Student double-booking detection
├── app/
│   ├── state.py                 # AppState dataclass — single source of truth
│   ├── events.py                # EventBus + event constants
│   ├── controller.py            # AppController — no tkinter
│   └── cost_config.py           # CostConfig dataclass — scheduler weights
├── file_io/
│   ├── timetable_reader.py      # Wraps build_timetable_tree_from_file
│   ├── state_repo.py            # StateRepository save/load
│   └── export.py                # to_pdf() and to_txt()
└── ui/
    ├── app.py                   # Thin shell — under 80 lines
    ├── constants.py             # UI colour and layout constants
    └── tabs/
        ├── timetable.py         # TimetableTab
        ├── verification.py      # VerificationTab
        └── exam.py              # ExamTab
```

### Key data models

| Class | Module | Role |
|---|---|---|
| `TimetableTree` | `core/timetable_tree.py` | Block → SubBlock → ClassList → StudentList |
| `ExamTree` | `reader/exam_tree.py` | GradeNode → ExamSubject → ClassList |
| `ExamPaper` | `reader/exam_paper.py` | One schedulable exam event |
| `ExamPaperRegistry` | `reader/exam_paper.py` | All papers; difficulty; links; constraints |
| `CostConfig` | `app/cost_config.py` | Scheduler penalty weights |
| `ScheduleResult` | `reader/exam_scheduler.py` | Dated slot assignments + student cost |

### Label formats

**Class labels:** `SUBJECT_TEACHERCODE_GRADE` — e.g. `AF_BALAY_08`, `MA_ALLEN_10`

**Paper labels:** `SUBJECT_P{n}_GRADE` — e.g. `MA_P1_Gr12`, `EN_P2_Gr10`

**Study paper labels:** `ST_P{n}_Gr{grade}` — e.g. `ST_P1_Gr12`

### State file format

```json
{
  "timetable_tree": {"A": {"A1": {"AF_BALAY_08": [361, 370]}}},
  "exclusions": ["LIB", "PE", "RDI"],
  "cost_config": {
    "same_week_penalty": 1,
    "teacher_load_penalty": 1,
    "day_density_factor": 5,
    "week_density_base": 6,
    "enforce_student_clash": true,
    "enforce_constraint_code": true,
    "teacher_tolerance_pct": 0
  },
  "papers": {
    "MA_12": {
      "papers": ["P1", "P2", "P3"],
      "constraints": ["C6"],
      "difficulty": "red",
      "links": [],
      "pinned_slot": null
    }
  },
  "session": {"start": "2026-10-12", "end": "2026-11-06", "am": true, "pm": true}
}
```

---

## Backlog

- Per-student printable exam timetable
- Simulated annealing timetable optimiser
- Phase 2: web UI via FastAPI