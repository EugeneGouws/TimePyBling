\# /verify



Run the full verification suite against the loaded timetable.



Steps:

1\. Check whether `tests/` exists and contains test files (directory listing only, no shell command).

&#x20;  - If test files exist, run `python -m pytest tests/ -q` and report results.

&#x20;  - If no test files exist, skip and continue.

2\. Run the clash check:

&#x20;  ```

&#x20;  python -c "from file\_io.state\_repo import StateRepository; from reader.verify\_timetable import find\_student\_clashes; s = StateRepository().load('data/timetable\_state.json'); clashes = find\_student\_clashes(s.timetable\_tree); print(f'Clashes: {len(clashes)}')"

&#x20;  ```

3\. Report any student double-bookings found.

4\. Report any test failures with the specific file and line number.

5\. If everything passes, confirm PASS with a one-line summary.

