Generate the exam schedule from the current state file and report results.

Steps:
1. Load data/timetable_state.json via StateRepository
2. Rebuild ExamTree and ExamPaperRegistry from loaded state
3. Call build_schedule() with the session config from state
4. Print a summary table:
   - Total papers scheduled
   - Total slots used
   - Student cost score
   - Number of warnings
5. List any warnings (same-week clashes, teacher marking conflicts)
6. List any grades where a student has 4+ exams in one week (hell week)
7. Do not write any files — display only unless user asks to export

Rules:
- Do not modify the scheduler to fix warnings — report them as-is
- If state file is missing, tell the user to run the app and Save State first
