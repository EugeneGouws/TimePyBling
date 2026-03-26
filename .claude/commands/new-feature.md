\# /new-feature



Scaffold a new feature following the TimePyBling layer architecture.



Ask the user:

1\. What does this feature do? (one sentence)

2\. Which layer does the core logic belong in? (core / reader / app / file\_io)

3\. Does it need a new UI element, or does it plug into an existing tab?



Then:

1\. Create the implementation file in the correct layer

2\. Add a stub with docstring and type hints — no implementation yet

3\. If `tests/` exists, add a corresponding empty test file in tests/

&#x20;  If `tests/` does not exist, skip the test file and note it in your response

4\. Add an import to the relevant \_\_init\_\_.py if needed

5\. Show the user the scaffolded files and confirm before writing any logic



Rules:

\- Never put logic in the wrong layer — check CLAUDE.md boundaries first

\- If the feature touches exam\_scheduler.py, flag it and ask before proceeding

