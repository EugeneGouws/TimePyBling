# HANDOFF.md — TimePyBling

Last updated: 2026-03-27
Branch: v2.0

## What was accomplished this session (Session 5)

**Two-weight cost function (complete):**
- core/cost_function.py — ColourWeight, StudentStressCost, TeacherMarkingCost,
  TotalCost. Convolution passes (window sizes 2/3/4). compute_windows() for
  efficient delta computation in hill-climb.
- Difficulty looked up from ExamPaperRegistry (ExamPaper has no colour field).
- Teacher-to-student-count precomputed from ExamTree class labels in constructor.

**Slot-swap hill-climb optimiser (complete):**
- core/hill_climb.py — hill_climb(schedule, cost_fn, max_iterations=10_000).
- Schedule format: {(day: int, session: str): list[ExamPaper]}.
- Swaps paper lists between slots; skips pinned papers; rejects swaps violating
  AM-before-PM constraint; delta computed over affected windows only.

**Controller integration (complete):**
- app/controller.py — optimise_schedule(stress_weight, marking_weight) method.
- Module-level helpers _build_schedule_dict() and _rebuild_schedule_result()
  convert between ScheduleResult and the hill-climb schedule dict format.

**UI integration (complete):**
- ui/constants.py — WEIGHT_STUDENT_STRESS = 50, WEIGHT_TEACHER_MARKING = 50.
- ui/tabs/exam.py — "Optimise" button added to cost panel; _on_optimise() handler
  reads slider weights, calls controller.optimise_schedule(), displays
  "Optimised — cost {old:.0f} → {new:.0f}" in the cost result label.

**Bash rule enforced:**
- CLAUDE.md updated: all bash commands must be single-line only.
- Memory updated with same rule.

**AM-before-PM plan agreed (not yet implemented — Priority 1 next session):**
- Full plan documented in next-session section below.

## Current state

| Layer     | Status                                           |
|-----------|--------------------------------------------------|
| core/     | stable + cost_function.py + hill_climb.py added  |
| reader/   | stable, Phase B + PenaltyEntry                   |
| app/      | built, optimise_schedule wired                   |
| file_io/  | built, new paper format                          |
| ui/       | Phase C complete + Optimise button               |
| tests/    | not yet built                                    |
| web/      | placeholder only                                 |

## Known bugs — carry forward from Session 4

1. Penalty breakdown not visible after schedule generation.
   Investigate: Breakdown button activation, penalty_log rendering in the
   Treeview, and constraint list display in the cell popout.

2. Cost panel layout not matching spec.
   Should show Hard Constraints section, Soft Constraints with editable
   weights, Optimisation Penalties with editable values, Rebuild button,
   and Breakdown button. Verify against Phase C spec.

3. Navigate-to-cell unverified.
   Confirm schedule popout clickable subject codes scroll the matrix canvas
   and flash the correct cell. Not yet verified.

4. State file session round-trip unverified.
   Exam session dates (start, end, AM/PM toggles) and slot assignments
   must persist in timetable_state.json on Save State and restore correctly
   on Load State.

## Next session plan (Session 6)

1. **AM-before-PM hard constraint (PRIORITY 1)** — plan agreed, implement first.
   Do NOT start anything else until this is done.

2. **OR-Tools CP-SAT optimiser (try after AM fix)** — replace or complement the
   hill-climb in core/hill_climb.py with a CP-SAT constraint programming solver
   to reach global optimum. AM-before-PM, difficulty, and clash constraints
   encoded as hard constraints; student stress and teacher marking load as the
   minimisation objective.

3. Fix known bugs 1-4 (priority order after items 1-2)
4. Manual verification pass: load Excel, configure papers, generate schedule,
   verify all popouts and round-trip save/load
5. Address any spec mismatches found during verification

### AM-before-PM plan (agreed 2026-03-27)

Files: reader/exam_scheduler.py (explicit instruction overrides stable-file rule),
       core/hill_climb.py

A. New helper _difficulty_order(registry, paper) -> int  red=0 yellow=1 green=2

B. New helper _am_first_slots(valid_slots, occupied_slots) -> list[int]
   AM slots (even index) always included. PM slot s only if s-1 in occupied_slots.
   Fallback to full valid_slots only if filtered list is empty.
   Applied before every _pick_spread_slot call.

C. Modified _pick_spread_slot score tuple: (s % 2, -load[s], gap)
   AM (s%2==0) is primary sort axis on top of the hard filter.

D. Replace Steps 1+2 with one sorted pass ordered by:
   (-grade_int, difficulty_order, -student_count)
   MA/PH still added to priority_labels for Step 3 compatibility.

E. Step 3 spacing pass: add _am_first_slots filter to candidate slots.

F. Step 5 hill-climb — two new guards after del assignment[lbl]:
   - Destination PM: skip if sl_new%2==1 and (sl_new-1) not in occupied
   - Source AM anchor: skip if sl_old%2==0 and sl_old not in occupied
     and sl_old+1 < num_slots and sl_old+1 in occupied

G. core/hill_climb.py: new _am_constraint_ok(sched, d1, d2) -> bool
   After swap: for each affected day, PM must not have papers without AM.
   Revert immediately if constraint violated, before computing cost delta.

## Backlog features (do not implement without instruction)

- Timetable tab: right-click empty subblock -> context menu to add LIB/ST/BAT
- Per-student printable exam timetable
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
  reader/exam_scheduler.py   <- explicit instruction pending for Session 6
  reader/exam_clash.py
  reader/exam_tree.py

Active UI file: ui/tabs/exam.py (NOT ui/ui.py -- that file is deleted)
```
