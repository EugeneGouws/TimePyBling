# CLAUDE.md — TimePyBling

Desktop tool: SA school timetable analysis + exam scheduling.
Stack: Python 3.13, Tkinter, pandas, openpyxl, reportlab (optional).
Target: Windows desktop app (pyinstaller), web-ready architecture (Phase 2).

## Commands

```sh
python main.py                   # launch app
pyinstaller TimePyBling.spec     # build Windows EXE
pip install pandas openpyxl reportlab
```

## Layer boundaries (never violate)

```
core/      pure domain — no UI, no I/O
reader/    pure scheduling logic — no UI, no I/O
app/       orchestration — no tkinter, no file I/O
file_io/   file I/O only — no tkinter, no business logic
ui/        tkinter only — no business logic
tests/     pytest only — no tkinter
web/       placeholder — do not add code here
```

Dependency direction: ui → app → core/reader, file_io → core/reader.
Nothing in core/ or reader/ imports from app/, file_io/, ui/, or web/.

## Stable files — do not modify without explicit instruction

- core/timetable_tree.py
- core/conflict_matrix.py
- reader/exam_clash.py
- reader/exam_tree.py

## Data flow

1. Excel (ST1.xlsx) → TimetableTree (one-time) → data/timetable_state.json
2. TimetableTree → ExamTree → ExamPaperRegistry
3. build_schedule() → ScheduleResult
4. Display in UI / export PDF or TXT

## Scheduling algorithm — 3-phase constructive (reader/exam_scheduler.py)

Step 0   Pinned papers (ST + any pinned_slot) placed first.
         ST papers reserve their slot for their grade.

Phase 1  RED + LINKED (AM only, 5-day spacing per grade, Gr12 → Gr08).
         Linked subjects auto-promoted to red.
         Primary paper placed in best AM slot; partner auto-placed in PM
         of the same day (fallback: best available AM slot).

Phase 2  YELLOW — fill gaps, AM first; PM unlocks only when AM exhausted.
         Most-constrained-first ordering. Minimises incremental per-student
         overlap cost at each placement. Grades Gr12 → Gr08.

Phase 3  GREEN — same logic as Phase 2.

Post     Paper-move hill-climb (20 passes) to locally minimise student cost.

Cost function: per-student overlap cost. Only penalises when a student has
2+ exams in the same convolution window (_PASSES = [(2,3),(3,2),(4,1)]).
Red papers 5+ days apart → zero stress score by design.

Teacher optimisation: separate Optimise step (hill_climb_teacher) after
Generate. Tolerance slider (0%=pure student, 100%=pure teacher) controls
how much student cost can increase to reduce teacher marking load.

## Excel input format

File: data/ST1.xlsx — one row per student.
Timetable columns: regex `^[A-H][1-9]$` (single digit only).
Cell format: `SUBJECT TEACHERCODE` e.g. `AF BALAY`.
Two-digit columns (F25–F67) are exam columns — excluded by regex.

## Naming conventions

- Class labels: `SUBJECT_TEACHERCODE_GRADE` e.g. `AF_BALAY_08`
- Paper labels: `SUBJECT_P{n}_GRADE` e.g. `MA_P1_Gr12`
- Grade labels: `Gr{nn}` e.g. `Gr08`, `Gr12`
- Study paper labels: `ST_P{n}_Gr{grade}` — counter derived from existing ST papers per grade
- UI methods: `_build_*` construct, `_on_*` handle events, `_refresh_*` update display
- Never put business logic inside a `_on_*` handler — delegate to controller

## Hard rules

- No tkinter imports outside ui/
- No pandas/openpyxl imports outside file_io/timetable_reader.py
- No bare except — catch specific exceptions
- No silent exception swallowing
- Single-line git commit messages only — no multi-line -m strings

## UI Scrolling Convention (Tkinter)

Every UI element that has a scrollbar must bind mouse scroll events:

Vertical scroll: <MouseWheel> → scrolls vertically (yview_scroll)
Horizontal scroll: <Shift-MouseWheel> → scrolls horizontally (xview_scroll)

This applies to all Treeview, Canvas, Listbox, Text, and Frame-with-scrollbar widgets. Bind on the widget itself and any child widgets that capture focus.

## Test infrastructure

tests/ does not exist yet. This is intentional.

- Do not check for tests during planning
- Do not reference missing tests in plans or summaries
- Do not suggest adding tests as part of any feature plan
- Do not run pytest or shell commands to check for tests
- When tests/ is built it will be announced explicitly

Until then, treat the absence of tests as a known project state, not a gap to flag.

## State file schema (current format)

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
    },
    "ST_1_12": {
      "papers": ["P1"],
      "constraints": [],
      "difficulty": "green",
      "links": [],
      "pinned_slot": 4
    }
  },
  "session": {"start": "2026-10-12", "end": "2026-11-06", "am": true, "pm": true}
}
```
Old format (papers as plain list) is auto-migrated on next save.

## Commit message rule (enforced)

Single-line only. Never use multi-line -m strings — triggers security warning.

```
git commit -m "type(scope): short description"
```

Never do this:
```
git commit -m "subject line
# this triggers the warning"
```

## Bash command rule (enforced)

Always write bash commands as single lines — no multi-line strings, no line
continuations, no heredocs. Break complex commands with && on one line.

```sh
# OK
python -c "import ast; ast.parse(open('f.py').read()); print('OK')"

# NEVER — causes shell break-character errors
python -c "
import ast
ast.parse(open('f.py').read())
"
```

## Mistakes log — update after every mistake

2026-03-28 — Session 7: Clean session, no mistakes.

2026-03-27 — Session 6: Static dash-style dividers in verification.py written as
             hardcoded `"─" * 60` etc. instead of matching the length of the
             adjacent text. Rule: always use len(heading_text) when writing
             dash dividers above/below headings in scrolled text widgets.
             Never use arbitrary round numbers like 60.

2026-03-27 — Session 5: Clean session, no mistakes.

2026-03-26 — Session 3: Edited ui/ui.py (legacy dead file, 108KB) instead of
             ui/tabs/exam.py. HANDOFF.md clearly states ui/ui.py was deleted
             in Session 1. Always check main.py import chain and HANDOFF.md
             before modifying UI files. Deleted ui/ui.py for real.

2026-03-26 — Column names hardcoded as "Name"/"Surname" in timetable_tree.py
             but actual Excel file has "SFirstname"/"SSurname". Student names
             weren't loading. Fixed column mapping in core/timetable_tree.py.

2026-03-26 — Session 4: Clean session, no mistakes.

2026-03-25 — Named file I/O layer io/ which clashes with Python stdlib io
             module (pre-loaded at startup). Renamed to file_io/. Never use
             stdlib module names for package folders.

## Known bugs — fix in next session

- FIXED 2026-03-26: Constraint codes not saving — from_exam_tree() wiped
  registry on rebuild. Fixed by adding prior parameter.
- FIXED 2026-03-26: Removing exclusion reset P2/P3 to P1 — same fix.
- FIXED 2026-03-28: Cost panel redesigned — old dual sliders replaced with
  single tolerance slider. Panel now matches current design intent.
- Penalty breakdown popout exists but penalty_log is always empty with the
  new 3-phase scheduler (PenaltyEntry never populated). Either populate it
  with per-student overlap data or remove the Breakdown button.
- Navigate-to-cell unverified. Confirm schedule popout subject codes scroll
  matrix canvas and flash correct cell.

## Backlog features — do not implement without instruction

- Per-student and per-teacher printable exam timetable
- Phase 2: web UI via FastAPI (web/ folder is placeholder)