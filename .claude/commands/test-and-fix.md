Run the test suite and fix any failures.

Steps:
1. Run `python -m pytest tests/ -v`
2. For each failing test, identify the root cause
3. Fix the failure — in the implementation, not by weakening the test
4. Re-run the affected test to confirm it passes
5. Run the full suite again to confirm no regressions
6. Report a summary: N passed, N fixed, 0 failing

Rules:
- Never delete or skip a test to make it pass
- Never catch exceptions in the implementation just to silence a test
- If a fix requires changing a stable file (core/ or reader/exam_scheduler.py), stop and ask the user
