# HANDOFF.md — TimePyBling

Last updated: 2026-03-28
Branch: v2.0

## What was accomplished this session (Session 7)

**Complete exam scheduler rewrite (complete):**
- reader/exam_scheduler.py — ground-up 3-phase constructive algorithm replacing
  all 7 old passes. Phase 1: red + linked papers, AM only, 5-day spacing,
  Gr12→Gr08. Phase 2: yellow, AM first, PM unlocks when AM exhausted. Phase 3:
  green, same logic. Post: paper-move hill-climb (20 passes).
- All warnings removed (ScheduledPaper.warnings, ScheduleResult.pin_clash_warnings,
  ScheduleResult.teacher_warnings).
- Dead helpers removed: _week_number, _label_to_parts, _difficulty_order,
  linked_pairs, linked_placed, _effective_order.

**Per-student overlap cost function (complete):**
- core/cost_function.py — StudentStressCost rewritten. Only penalises when a
  specific student has 2+ exams in the same convolution window. Red papers
  5+ days apart → guaranteed zero stress.
- TeacherMarkingCost gains position ramp: (window_start+1)/total_days.

**Tolerance-based teacher optimisation (complete):**
- core/hill_climb.py — hill_climb_teacher() added. Accepts swaps only if teacher
  cost decreases AND student cost stays within baseline × (1+tolerance/100).
- core/cpsat_optimiser.py — complete rewrite: cpsat_optimise_teacher() minimises
  teacher cost subject to student cost upper bound. Retained for future use.
- app/controller.py — generate_schedule() now saves student_optimal_result as
  baseline; optimise_schedule(teacher_tolerance_pct) returns 5-tuple.

**UI overhaul (complete):**
- ui/constants.py — dual weight constants replaced by DEFAULT_TEACHER_TOLERANCE=0.
- ui/tabs/exam.py — dual sliders replaced by single tolerance slider
  (0%=pure student, 100%=pure teacher). All warning display removed.
  _on_optimise() updated to new 5-return-value signature.

**State migration (complete):**
- file_io/state_repo.py — strips old student_stress_weight/teacher_load_weight
  from saved state files before constructing CostConfig.

**AppState (complete):**
- app/state.py — student_optimal_result field added.
- app/cost_config.py — teacher_tolerance_pct added; old weight fields removed.

## Current state

| Layer    | Status                                                        |
|----------|---------------------------------------------------------------|
| core/    | ✅ stable — cost_function, hill_climb, cpsat_optimiser all done |
| reader/  | ✅ 3-phase scheduler complete, no warnings                     |
| app/     | ✅ generate + optimise wired, student_optimal_result baseline  |
| file_io/ | ✅ state migration for old cost_config format                  |
| ui/      | ✅ tolerance slider, warning-free render                       |
| tests/   | ❌ not yet built                                               |
| web/     | ❌ placeholder only                                            |

## Known bugs

1. Penalty breakdown popout exists but penalty_log is always empty with the new
   scheduler (PenaltyEntry never populated). Either remove the Breakdown button
   or populate penalty_log with per-student overlap entries.

2. Navigate-to-cell unverified. Confirm schedule popout clickable subject codes
   scroll the matrix canvas and flash the correct cell.

## Backlog features — do not implement without instruction

- Per-student and per-teacher printable exam timetable
- Phase 2: web UI via FastAPI (web/ folder is placeholder)

## Next session plan (Session 8)

**Priority issues flagged by user after session end:**

1. **AM-first mechanic needs per-grade AM slot tracking** — the current
   `pm_unlocked` flag is global across all grades. Phases 2 and 3 need to
   track remaining AM slots per grade independently so AM is exhausted
   per-grade before PM unlocks for that grade.

2. **Hill-climb needs more passes** — 20 passes is not enough. Verify via the
   debug window and increase until convergence. Consider a convergence check
   rather than a fixed pass count.

3. **CostConfig needs updating** — legacy fields (same_week_penalty,
   teacher_load_penalty, day_density_factor, week_density_base) are dead
   weight. Review and strip fields that no longer drive the new algorithm.

4. **5-day red paper spacing needs tweaking** — the current hard 5-day gap may
   be too rigid or not working correctly in all cases. Debug and adjust.

5. **Fix or remove Breakdown button** — penalty_log is always empty. Either
   remove the button, or populate penalty_log from the per-student overlap data
   after placement.

## Architecture reminder

Dependency direction: ui → app → core/reader, file_io → core/reader.
Nothing in core/ or reader/ imports from app/, file_io/, ui/, or web/.

Stable files (do not modify without explicit instruction):
- core/timetable_tree.py
- core/conflict_matrix.py
- reader/exam_clash.py
- reader/exam_tree.py
