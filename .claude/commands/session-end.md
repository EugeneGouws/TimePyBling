\# /session-end



End-of-session wrap-up. Updates docs, commits, and pushes.

Every destructive step requires explicit approval before proceeding.



\---



\## Step 1 — Check for tests



Check whether `tests/` exists and contains any `test\_\*.py` files

by inspecting the directory listing only. Do not run any shell command.



If no test files exist: print "No tests — skipping" and move to Step 2.

Do not run pytest. Do not warn about missing tests. Do not suggest adding tests.



If test files exist, run:

```

pytest tests/ -v --tb=short

```

Show full output. If any tests fail, STOP and report failures.

Do not proceed to Step 2 until all tests pass or the user types "skip tests".



\---



\## Step 2 — Summarise session changes



List every file modified, created, or deleted this session with a

one-line description of what changed. Group by layer:

core/, reader/, app/, file\_io/, ui/, tests/, docs



Ask: "Does this summary look correct? Type 'yes' to continue or describe corrections."



Wait for explicit approval before proceeding.



\---



\## Step 3 — Update CLAUDE.md



Open CLAUDE.md. Add an entry to the mistakes log for any errors made

this session (wrong file edited, incorrect assumption, bug introduced

and fixed). If no mistakes were made, add a line confirming the session was clean.



Show the proposed addition. Ask: "Approve CLAUDE.md update? (yes/edit/skip)"



Wait for explicit approval before writing.



\---



\## Step 4 — Update README.md



Review the current README against the session changes. Update:

\- Project structure section if files were added or removed

\- Feature descriptions if new functionality was built

\- Any outdated instructions or state file format examples



Show a diff of proposed changes. Ask: "Approve README update? (yes/edit/skip)"



Wait for explicit approval before writing.



\---



\## Step 5 — Update HANDOFF.md



Rewrite HANDOFF.md with the following sections:



\*\*Last updated:\*\* today's date and branch name



\*\*What was accomplished this session:\*\* bullet list of completed work



\*\*Current state table:\*\* all layers with status (✅ stable / ⚠️ in progress / ❌ not started)



\*\*Known bugs:\*\* any new bugs discovered, and status of previously listed bugs



\*\*Backlog features:\*\* unchanged from previous HANDOFF unless items were completed or added



\*\*Next session plan:\*\* priority-ordered list of what to do next, with the top item clearly marked



\*\*Architecture reminder:\*\* dependency direction, stable files list



Show the full proposed HANDOFF.md. Ask: "Approve HANDOFF update? (yes/edit/skip)"



Wait for explicit approval before writing.



\---



\## Step 6 — Stage and commit



Run:

```

git status

git diff --stat

```



Show output. Propose a commit message in this format:

```

type(scope): short description

```

Where type is one of: feat / fix / refactor / docs / test / chore



Ask: "Approve this commit message? (yes/edit/skip)"



Wait for explicit approval. If approved, run:

```

git add -A

git commit -m "type(scope): short description"

```



Single-line commit message only. No multi-line -m strings. No newlines inside quotes.



\---



\## Step 7 — Push



Show the target remote and branch:

```

git remote -v

git branch --show-current

```



Ask: "Push to origin/<branch>? (yes/skip)"



Wait for explicit approval. If approved, run:

```

git push origin <current-branch>

```



Report success or failure. If push fails, report the error and stop.

Do not force push under any circumstances.



\---



\## Completion



Report a one-line summary: tests status, files updated, commit hash, push status.

