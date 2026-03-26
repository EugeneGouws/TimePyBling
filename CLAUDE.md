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
- reader/exam_scheduler.py
- reader/exam_clash.py
- reader/exam_tree.py

## Data flow

1. Excel (ST1.xlsx) → TimetableTree (one-time) → data/timetable_state.json
2. TimetableTree → ExamTree → ExamPaperRegistry
3. build_schedule() → ScheduleResult
4. Display in UI / export PDF or TXT

## Scheduling algorithm — 7 passes in order (do not reorder)

0. ST pinned papers placed first — reserves grade_reserved_slots
0.5. Linked paper pre-placement — pairs placed consecutively before DSatur
1. Priority papers (MA, PH) placed Gr12→Gr08 with max spread
2. Remaining papers via DSatur saturation ordering
3. Spacing pass — same-subject same-grade same ISO week flagged/swapped
4. Teacher marking load pass — same subject + teacher in same slot flagged
5. Hill-climb — quadratic penalty: day=k*(k-1)*day_density_factor,
   week=(k-1)*(k+week_density_base)//2. AM-first tiebreak on equal cost moves.

Difficulty enforcement: _difficulty_allows() checked before every placement.
Red blocks Red+Yellow same day. Yellow blocks Red same day. Green never blocks.
Linked partners are exempt from difficulty clash checks with each other.

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
    "enforce_constraint_code": true
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

## Mistakes log — update after every mistake

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
- Penalty breakdown and constraint list not visible after schedule generation.
  Investigate Breakdown button activation, penalty_log rendering, and
  constraint list in cell popout.
- Cost panel layout not matching spec. Should show Hard Constraints, Soft
  Constraints with editable weights, Optimisation Penalties, Rebuild and
  Breakdown buttons. Verify against Phase C spec.
- Navigate-to-cell unverified. Confirm schedule popout subject codes scroll
  matrix canvas and flash correct cell.
- State file session round-trip unverified. Session dates, AM/PM toggles,
  and slot assignments must persist on Save/Load State.

## Backlog features — do not implement without instruction

- COMPLETED 2026-03-26: Timetable tab: clicking a class shows student names.
- Timetable tab: right-click empty subblock → context menu to add LIB/ST/BAT
- Per-student printable exam timetable
- Simulated annealing optimiser
- Phase 2: web UI via FastAPI (web/ folder is placeholder)
- Post-Phase C: consecutive teacher marking penalty
- Post-Phase C: end-of-timetable clustering penalty