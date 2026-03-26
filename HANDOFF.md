# HANDOFF.md — TimePyBling

Last updated: 2026-03-26
Branch: v2.0

## What was accomplished this session (Session 2)

**Cleanup & Polish:**
- ✓ Removed abandoned git worktree (lucid-kowalevski)
- ✓ Deleted stale remote branch (origin/master)
- ✓ Removed build/ and dist/ artifacts from tracking
- ✓ Cleaned __pycache__ directories

**Timetable Tab Enhancement:**
- ✓ Added expandable popout showing classes + student names
- ✓ Two-column layout: classes (left) → click class → students (right)
- ✓ Student names now display correctly (fixed column mapping: Name/Surname → SFirstname/SSurname)
- ✓ Removed unused "View Timetable" button

**Bug Fixes:**
- ✓ Fixed student name column mapping in core/timetable_tree.py (481 names now load)

**Documentation:**
- ✓ Updated CLAUDE.md mistakes log
- ✓ Marked first backlog feature as completed
- ✓ Updated this handoff with exam tab priorities

## What was built Session 1 (Phase 1 refactor)

Full layered architecture refactor. App runs identically to V1.0.

### New files
```
CLAUDE.md                        — gremlin's project instructions
.claude/commands/verify.md       — slash command
.claude/commands/test-and-fix.md — slash command
.claude/commands/commit.md       — slash command
.claude/commands/generate.md     — slash command
.claude/commands/new-feature.md  — slash command
app/__init__.py
app/state.py                     — AppState dataclass, single source of truth
app/events.py                    — EventBus + event constants
app/controller.py                — AppController, no tkinter
file_io/__init__.py
file_io/timetable_reader.py      — wraps build_timetable_tree_from_file
file_io/state_repo.py            — StateRepository save/load/apply_pending_papers
file_io/export.py                — to_pdf() and to_txt()
ui/constants.py                  — SESSIONS = ["AM", "PM"] (ui-layer copy)
ui/tabs/__init__.py
ui/tabs/timetable.py             — TimetableTab
ui/tabs/verification.py          — VerificationTab
ui/tabs/exam.py                  — ExamTab
ui/app.py                        — thin shell, under 80 lines
```

### Deleted files
```
ui/ui.py                         — replaced by ui/app.py + ui/tabs/
```

### Key decisions
- `io/` renamed to `file_io/` — Python stdlib `io` is pre-loaded at startup
- `ui/` imports from `app/` only — never from `core/` or `reader/` directly
- Duplicate `SESSIONS` constant in `ui/constants.py` is intentional — layer boundary takes precedence over DRY for trivial values
- `app/controller.py` uses lazy imports for `file_io/` to keep testing clean

## Current state

| Layer | Status |
|---|---|
| core/ | ✅ unchanged, stable |
| reader/ | ✅ unchanged, stable |
| app/ | ✅ built |
| file_io/ | ✅ built |
| ui/ | ✅ refactored |
| tests/ | ❌ not yet built |
| web/ | ❌ placeholder only |

## Known bugs (fix next session)

**Priority 1 — Exam Tab Bug Fixes (critical user-facing issues):**
- Exam tab: manual constraint codes not saving via constraint entry field
- Exam tab: removing a subject exclusion (e.g. MU_08) resets added P2/P3 papers back to P1

Both suspected root cause: ExamTab event handling + AppController._rebuild_exam()
wiping ExamPaperRegistry without preserving added papers.

**Files to examine:** ui/tabs/exam.py, app/controller.py

## Backlog features (do not implement without instruction)

- Timetable tab: click a class node → show student list in detail panel
- Timetable tab: right-click empty subblock → context menu to add LIB/ST/BAT
  (ST = students only, BAT = teachers only, LIB = both)
- Per-student printable exam timetable
- Simulated annealing optimiser
- Phase 2: web UI via FastAPI (web/ folder is placeholder)

## Next session plan

**Session 3 (Priority order):**

1. **FIX EXAM TAB BUGS** (top priority)
   - Constraint codes not saving
   - Exclusion removal resets papers
   - Root cause: _rebuild_exam() wiping state

2. Build tests/ — pytest for domain logic only, no tkinter
   - Add basic integration tests for exam_scheduler

3. Pick a backlog feature
   - Right-click context menu for LIB/ST/BAT (if time)

**Session 3 session notes:**
- Exam tab is the highest-priority user-facing component
- Both bugs block normal user workflows
- Tests infrastructure needed before more changes

## Architecture reminder

```
Dependency direction: ui → app → core/reader, file_io → core/reader
Nothing in core/ or reader/ imports from app/, file_io/, ui/, or web/
Stable files: core/timetable_tree.py, core/conflict_matrix.py, reader/exam_scheduler.py
```
