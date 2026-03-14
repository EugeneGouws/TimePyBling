# TimePyBling

School timetable analysis tool. Reads a completed student timetable, checks it for clashes, scores it against a cost function, and schedules exams.

## Requirements

- Python 3.13
- `pip install pandas openpyxl`

## Data files

Place these in the `data/` folder before running:

| File | Description |
|---|---|
| `data/ST1.xlsx` | Student timetable — one row per student, columns A1–H7 |
| `data/teachers.xlsx` | Teacher codes and subject pools |

ST1 cell format: `SUBJECT TEACHERCODE` e.g. `AF BALAY`. Empty slots can be left blank or filled with `FREE`.

## Running

```
python main.py
```

This launches the UI. Load files using the buttons in the top bar.

## Project structure

```
TimePyBling/
├── main.py                    # Entry point — launches the UI
├── data/
│   ├── ST1.xlsx
│   └── teachers.xlsx
├── core/
│   ├── timetable_tree.py      # Parse ST1.xlsx → TimetableTree (read-only)
│   ├── block_tree.py          # Mutable BlockTree for SA moves
│   └── conflict_matrix.py     # Pure shared-student conflict utility
├── reader/
│   ├── exam_tree.py           # Reorganise TimetableTree by grade + subject
│   ├── exam_clash.py          # DSatur exam slot scheduling
│   └── verify_timetable.py    # Detect student and teacher double-bookings
├── optimizer/
│   ├── cost_function.py       # E(T) cost evaluator
│   └── block_exporter.py      # BlockTree → ST1.xlsx output
└── ui/
    └── ui.py                  # tkinter interface
```

## UI tabs

**Timetable** — browse the full timetable tree (Block → SubBlock → Class → Students). Live search by student ID, subject code, or teacher name.

**Verification** — loads automatically on file open. Shows:
- Clash report: student and teacher double-bookings
- Cost function E(T): per-term breakdown
- Teacher qualifications: cross-checks teacher assignments against their subject pool in teachers.xlsx

**Exams** — exam slot scheduling by grade using DSatur graph colouring. Manage subject exclusions (ST, LIB, PE, RDI excluded by default). Rebuild as needed.

**Export** — write the timetable back to ST1.xlsx. Requires `timetable_converter.py` (not yet written — see below).

## Cost function E(T)

| Term | Description                                    | Weight |
|---|------------------------------------------------|---|
| C_s | Student double-bookings                        | 10 000 |
| C_t | Teacher double-bookings                        | 5 000 |
| P_g12 | Gr 12 class taught by non-preferred teacher    | 400 |
| P_tg | Teacher assigned outside preferred grades      | 100 |
| P_f | Teacher has no free on a specific day          | 50 |
| P_stg | Sparse subject group on consecutive cycle days | 10 |
| P_alloc | Allocation rules Gr 9–11                       | 5 |

P_g12, P_tg, P_f, and P_alloc are stubs — they return 0 until the relevant optional columns are added to teachers.xlsx.

## teachers.xlsx format

Required columns:

| Column | Description |
|---|---|
| `Teacher Code` | Code used in ST1 cells e.g. `BALAY` |
| `sua` | Primary subject code |
| `sub` | Secondary subject code (optional) |
| `suc` | Tertiary subject code (optional) |

Optional columns to activate preference terms:

| Column | Description                                           |
|---|-------------------------------------------------------|
| `gr12` | Each Gr 12 learner gets to choose one subject teacher |
| `pref_grades` | Comma-separated e.g. `10,11,12`                       |

## What is not yet built

`timetable_converter.py` — converts a TimetableTree into a BlockTree so the export and SA optimiser can run. This is the next piece to write. Once it exists, the Export tab and the simulated annealing optimiser can be wired up.

## Class label format

All class labels follow `SUBJECT_TEACHERCODE_GRADE`:

- `AF_BALAY_08` — subject AF, teacher BALAY, grade 8
- `MA_ALLEN_10` — subject MA, teacher ALLEN, grade 10
- `DR_VAN_DEN_BERG_12` — teacher codes with underscores are handled correctly

Grade is always zero-padded to two digits. The subject merge `OD → DR` is applied at parse time in `timetable_tree.py`.