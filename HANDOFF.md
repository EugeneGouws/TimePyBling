# HANDOFF.md — TimePyBling

Last updated: 2026-03-26
Branch: v2.0

## What was accomplished this session (Session 4)

**Pre-C fix (complete):**
- Wired CostConfig into build_schedule(). Replaced w_day/w_week params with
  single `config` param (duck-typed, no CostConfig import in reader/).
  Weights extracted at function top with hardcoded fallbacks.
- Step 3 spacing swap gated on `_w_same_week > 0`.
- Step 5 hill-climb closures use `_w_day`, `_w_week_base`, `_w_same_week`.
- AppController passes `config=self._state.cost_config` to build_schedule().

**Phase C — UI rewrite (complete):**
- Replaced Treeview exam tree with scrollable Canvas matrix (subjects x grades).
  Cells colored by difficulty (red/yellow/green), hover highlight, click-to-open
  cell popout.
- Cell popout: paper list, add/remove paper, difficulty radio buttons, links
  editor, constraint editor per paper, pin/unpin per paper. Auto-refreshes on
  mutation.
- ST row at bottom of matrix with per-grade study paper management popout.
- Session config controls (start/end dates, AM/PM, configure sessions) moved
  to left pane header bar.
- Cost function panel rewired to CostConfig fields (day_density_factor,
  week_density_base, same_week_penalty, teacher_load_penalty). Builds
  CostConfig from UI entries before calling generate_schedule().
- Display-only checkboxes for enforce_student_clash, enforce_constraint_code.
- Rebuild Schedule button in cost panel runs _generate_exam_schedule.

**C4b — Penalty breakdown (complete):**
- Added PenaltyEntry dataclass to reader/exam_scheduler.py.
- Added penalty_log: list[PenaltyEntry] to ScheduleResult.
- Scheduler accumulates day_density, week_density, teacher_load entries after
  hill-climb completes.
- Breakdown button opens Toplevel with sortable Treeview, summary line, and
  Export as text button.
- Cost display auto-updates from penalty_log after schedule generation.

**Controller additions:**
- set_difficulty(subject, grade, difficulty)
- add_link(label_a, label_b) / remove_link(label_a, label_b)
- add_study_paper(grade, pinned_slot) / remove_study_paper(label)
All delegate to registry and publish EVT_PAPERS_CHANGED.

**Color constants added to ui/constants.py:**
- CLR_DIFF_RED, CLR_DIFF_YELLOW, CLR_DIFF_GREEN, CLR_DIFF_GRAY
- CLR_HOVER, CLR_ST_ROW

**Dead code removed:**
- _populate_exam_tree, _exam_tree_get_state, _exam_tree_restore_state
- _on_exam_tree_select, _refresh_paper_panel, _refresh_paper_panel_multi
- _set_constraint_ui_enabled, _on_paper_select, _refresh_constraint_list
- _current_subject_grade, _calculate_exam_cost

## What was accomplished Session 3

**Phase A — Plumbing (complete):**
- Fixed both known bugs: from_exam_tree() accepts prior registry, copies
  constraints, difficulty, links, pinned slots, and P2/P3 papers across rebuilds.
  Ghost subject guard prevents copying orphaned entries.
- Added links: set[str] field to ExamPaper
- Added subject-level difficulty (_difficulty dict) to ExamPaperRegistry
  with get_difficulty() / set_difficulty()
- Created app/cost_config.py — CostConfig dataclass with soft weights
  and hard constraint toggles, wired into AppState
- Updated file_io/state_repo.py — new paper format (dict per entry),
  auto-migrates old list format, serialises cost_config
- Added add_study_paper() / remove_study_paper() to registry
- Deleted legacy ui/ui.py (dead code, replaced in Session 1)

**Phase B — Scheduler (complete):**
- AM-first tiebreak in hill-climb Step 5
- _difficulty_allows() in slot assignment loop — Red blocks Red+Yellow same
  day; Yellow blocks Red; Green never blocks; linked partners exempt
- Linked pre-placement pass (Step 0.5) before DSatur — pairs placed
  consecutively, fallback with warning if slot+1 blocked
- ST pinned papers pre-committed as grade_reserved_slots before all
  other placement (Step 0)
- All 7 smoke tests pass

## Current state

| Layer     | Status                              |
|-----------|-------------------------------------|
| core/     | stable, unchanged                   |
| reader/   | stable, Phase B + PenaltyEntry      |
| app/      | built, CostConfig wired, 5 new methods |
| file_io/  | built, new paper format             |
| ui/       | Phase C complete, known bugs below  |
| tests/    | not yet built                       |
| web/      | placeholder only                    |

## Known bugs — fix in next session

1. **Penalty breakdown not visible after schedule generation.**
   Investigate: Breakdown button activation, penalty_log rendering in the
   Treeview, and constraint list display in the cell popout.

2. **Cost panel layout not matching spec.**
   Should show Hard Constraints section, Soft Constraints with editable
   weights, Optimisation Penalties with editable values, Rebuild button,
   and Breakdown button. Verify against Phase C spec.

3. **Navigate-to-cell unverified.**
   Confirm schedule popout clickable subject codes scroll the matrix canvas
   and flash the correct cell. Not yet verified.

4. **State file session round-trip unverified.**
   Exam session dates (start, end, AM, PM toggles) and slot assignments
   must persist in timetable_state.json on Save State and restore correctly
   on Load State. Verify save and load round-trip for session config.

## Next session plan (Session 5)

1. Fix known bugs 1-4 above (priority order)
2. Manual verification pass: load Excel, configure papers, generate schedule,
   verify all popouts and round-trip save/load
3. Address any spec mismatches found during verification

## Backlog features (do not implement without instruction)

- Timetable tab: right-click empty subblock -> context menu to add LIB/ST/BAT
- Per-student printable exam timetable
- Simulated annealing optimiser
- Phase 2: web UI via FastAPI
- Post-Phase C: consecutive teacher marking penalty (soft constraint)
- Post-Phase C: end-of-timetable clustering penalty (soft constraint)

## Architecture reminder

```
Dependency direction: ui -> app -> core/reader, file_io -> core/reader
Nothing in core/ or reader/ imports from app/, file_io/, ui/, or web/

Stable files (do not modify without explicit instruction):
  core/timetable_tree.py
  core/conflict_matrix.py
  reader/exam_scheduler.py
  reader/exam_clash.py
  reader/exam_tree.py

Active UI file: ui/tabs/exam.py (NOT ui/ui.py -- that file is deleted)
```
