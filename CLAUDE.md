# CLAUDE.md — TimePyBling

Desktop tool: SA school timetable analysis + exam scheduling.
Stack: Python 3.13, Tkinter, pandas, openpyxl, reportlab (optional).
Target: Windows desktop app (pyinstaller), web-ready architecture (Phase 2).

## Commands

```sh
python main.py                   # launch app
python -m pytest tests/ -q       # run tests
pyinstaller TimePyBling.spec     # build Windows EXE
pip install pandas openpyxl reportlab
```

## Layer boundaries (never violate)

```
core/      pure domain — no UI, no I/O
reader/    pure scheduling logic — no UI, no I/O
app/       orchestration — no tkinter, no file I/O
file_io/        file I/O only — no tkinter, no business logic
ui/        tkinter only — no business logic
tests/     pytest only — no tkinter
web/       placeholder — do not add code here
```

Dependency direction: ui → app → core/reader, file_io → core/reader.
Nothing in core/ or reader/ imports from app/, file_io/, ui/, or web/.

## Stable files — do not modify without explicit instruction

- core/timetable_tree.py
- core/conflict_matrix.py
- reader/exam_scheduler.py

## Data flow

1. Excel (ST1.xlsx) → TimetableTree (one-time) → data/timetable_state.json
2. TimetableTree → ExamTree → ExamPaperRegistry
3. build_schedule() → ScheduleResult
4. Display in UI / export PDF or TXT

## Scheduling algorithm — 5 passes in order (do not reorder)

1. Pinned papers placed first
2. Priority papers (MA, PH) placed Gr12→Gr08 with max spread
3. Remaining papers via DSatur saturation ordering
4. Spacing pass — same-subject same-grade same ISO week flagged/swapped
5. Hill-climb — quadratic penalty: day=k*(k-1)*5, week=(k-1)*(k+6)//2

## Excel input format

File: data/ST1.xlsx — one row per student.
Timetable columns: regex `^[A-H][1-9]$` (single digit only).
Cell format: `SUBJECT TEACHERCODE` e.g. `AF BALAY`.
Two-digit columns (F25–F67) are exam columns — excluded by regex.

## Naming conventions

- Class labels: `SUBJECT_TEACHERCODE_GRADE` e.g. `AF_BALAY_08`
- Paper labels: `SUBJECT_P{n}_GRADE` e.g. `MA_P1_Gr12`
- Grade labels: `Gr{nn}` e.g. `Gr08`, `Gr12`
- UI methods: `_build_*` construct, `_on_*` handle events, `_refresh_*` update display
- Never put business logic inside a `_on_*` handler — delegate to controller

## Hard rules

- No tkinter imports outside ui/
- No pandas/openpyxl imports outside file_io/timetable_reader.py
- No bare except — catch specific exceptions
- No silent exception swallowing
- All tests must pass before committing

## State file schema (do not change)

```json
{
  "timetable_tree": {"A": {"A1": {"AF_BALAY_08": [361, 370]}}},
  "exclusions": ["LIB", "PE", "RDI", "ST"],
  "papers": {"MA_12": ["P1", "P2", "P3"]},
  "session": {"start": "2026-10-12", "end": "2026-11-06", "am": true, "pm": true}
}
```

## Mistakes log — update after every mistake

2026-03-25 — Named file I/O layer io/ which clashes with Python stdlib io module 
             (pre-loaded at startup). Renamed to file_io/. Never use stdlib module 
             names for package folders.

## Known bugs — fix in next session

- Exam tab: manual constraint codes are not saving when added via the 
  constraint entry field
- Exam tab: removing a subject exclusion (e.g. MU_08) resets any 
  previously added P2/P3 papers back to P1 only

## Backlog features — do not implement without instruction

- Timetable tab: clicking a class/lesson node in the timetable tree 
  should display the full student list for that class in a detail panel 
  or popout. Currently clicking does nothing.

- Timetable tab: right-click on an empty subblock cell should show a 
  context menu to add LIB, ST, or BAT to that slot.
  Rules:
    - ST can only be added to student rows
    - BAT can only be added to teacher rows  
    - LIB available to both
  This requires understanding whether a row is a student or teacher 
  context — factor that into the design.