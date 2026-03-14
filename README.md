# TimePyBling

A school timetabling and data analytics platform built in Python.

TimePyBling ingests student subject choices and teacher qualifications,
analyses scheduling constraints, generates conflict-free timetable blocks,
and exports back to the ST1 format used by the school database.

The long-term vision is a data and analytics framework for schools and
universities built around the timetable as the
core data structure.

---

## Project structure

```
TimePyBling/
├── data/                    ← Not committed — see Data files section below
├── output/                  ← Not committed — generated ST1 exports land here
│
├── xlsx_loader.py           ← Generic xlsx row parser, reused by all loaders
├── conflict_matrix.py       ← Pure conflict logic, context-free
│
├── school_tree.py           ← Student subject enrolment tree
├── teacher_tree.py          ← Teacher qualification tree
├── conflict_analyser.py     ← Timetable and exam conflict analysers
├── check_data.py            ← Data quality gate — run before scheduling
│
├── timetable_tree.py        ← Reads ST1.csv / ST1.xlsx into memory
├── exam_tree.py             ← Builds exam subject groups from timetable
├── exam_clash.py            ← DSatur + backtracking exam slot solver
├── main.py                  ← Entry point for exam clash workflow
│
├── block_tree.py            ← Core data structure for new timetable builder
├── block_builder.py         ← Constraint-based block assignment algorithm
├── block_exporter.py        ← Exports BlockTree to ST1-format xlsx
│
└── ui.py                    ← GUI interface
```

---

## Data files

Data files are not committed to this repository. Place the following
files in the `data/` folder before running any scripts.

### `data/students.xlsx`

One row per student. No header row required — detected automatically.

| Column | Content | Example |
|--------|---------|---------|
| 0 | Student ID (numeric) | `480` |
| 1 | Grade (numeric) | `8` |
| 2+ | Subject codes, one per column | `EN`, `MA`, `AF`, `SC` |

Subject codes are uppercase abbreviations matching the school's
subject register. Empty cells are ignored — students may have
different numbers of subjects.

### `data/teachers.xlsx`

One row per teacher. No header row required.

| Column | Content | Example |
|--------|---------|---------|
| 0 | Teacher ID (numeric) | `1` |
| 1 | Teacher code | `EG` |
| 2+ | Subject codes they can teach, one per column | `MU`, `MA` |

Teachers may teach up to as many subjects as there are columns.
Empty cells are ignored.

### `ST1.xlsx`

The current timetable export from the school database. Place in the
project root (not in `data/`). Used by the exam clash workflow.

| Column | Content |
|--------|---------|
| `Studentid` | Student ID (numeric) |
| `Grade` | Grade (numeric) |
| `A1`–`H7` | Class assignment per slot, e.g. `AF BALAY` |

---

## Subjects excluded from automated scheduling

The following subject codes are treated as manually allocated and
excluded from conflict analysis and block assignment:

| Code | Reason |
|------|--------|
| `LIB` | Library — allocated manually |
| `Study` | Free/study periods — allocated manually |
| `RDI` | Placed inside the EN block automatically |

---

## Quickstart

### 1. Check data quality

```bash
python check_data.py
```

Verifies all student subjects have at least one qualified teacher,
flags bottleneck subjects, and prints conflict matrices.

### 2. Run exam clash analysis

```bash
python main.py
```

Reads `ST1.xlsx` (or `ST1.csv`), builds the exam tree, and prints
the minimum exam slot report per grade.

### 3. Build and export a new timetable

```bash
python block_builder.py
python block_exporter.py
```

Builds a BlockTree from student and teacher data and exports it
to `output/ST1_new.xlsx` in ST1 format.

---

## Dependencies

```bash
pip install openpyxl pandas
```

Python 3.11 or later recommended (uses `X | Y` type union syntax).

---

## Algorithm overview

The block assignment problem is NP-hard (reducible to graph colouring).
The approach taken here mimics the manual process used by experienced
human schedulers:

1. **Precompute** legal subblock placement patterns per subject group
2. **Score** subject groups by constraint pressure — most constrained first
3. **Construct greedily** with forward checking to preserve future options
4. **Repair locally** using simulated annealing or tabu search

The conflict matrix is the foundation — two subject groups that share
at least one student cannot occupy the same slot. This is the same
logic used for exam clash detection, generalised into a reusable
`ConflictMatrix` class.

See `block_builder.py` for the full algorithm skeleton and TODO markers.

---

## Roadmap

- [ ] Stage 3: Greedy construction with forward checking
- [ ] Stage 4: Simulated annealing repair
- [ ] GUI for manual block editing (drag and drop)
- [ ] Teacher load balancing and optimisation
- [ ] Analytics layer — subject popularity, clash rates, teacher demand
- [ ] Multi-school / district support
