Company AI Excel Batch Tool V5 - Fixed 4 Concurrent / 2 Runs
============================================================

Purpose
-------
Read text from an Excel column, send each row text into the internal Company AI
by embedding it into the Prompt, and write the answers back to the configured
result columns on the same row.

Each row is processed TWICE (RUN_COUNT = 2). Run 1 goes to 结果列1, Run 2 goes to
结果列2, so the two outputs can be compared side by side.

V5 concurrency design
---------------------
- Fixed concurrency: 4 workers. It is not exposed in the UI.
- Up to 4 Excel rows can call the AI at the same time.
- The 2 runs of the same row are executed one after another, and each run opens
  its own AI session, so the two answers do not share conversation context.
- All requests share the Authorization Bearer token entered in the UI.
- Every row creates its own AI session, so conversation context is not shared.
- API workers never write Excel directly.
- Excel writes/saves are performed serially by the batch controller to prevent
  openpyxl/workbook write conflicts.
- If Stop is clicked, no new rows are submitted. The current maximum 4 in-flight
  requests are allowed to finish and are saved, so the file can be resumed later.

Fixed settings
--------------
- Model: V4 / ds_v4
- Agent ID: standard
- Thinking: False
- SSL verification: False
- Worker cooldown after each row: 1 second
- Runs per row: 2 (RUN_COUNT in the .py file)

Prompt behavior
---------------
- If the Prompt box is blank, DEFAULT_PROMPT in company_ai_excel_tool.py is used.
- If text is entered in the Prompt box, that text overrides DEFAULT_PROMPT for
  the current run.
- The employee prompt of the current row is inserted INTO the fixed Prompt, not
  appended after it. Substitution rules, in priority order:
    1. If the Prompt contains {employee_prompt}, that placeholder is replaced
       with the employee prompt (this is what DEFAULT_PROMPT uses).
    2. Else if the Prompt contains the pair
       <<<EMPLOYEE_PROMPT_START ... EMPLOYEE_PROMPT_END>>>,
       the text between the two markers is replaced with the employee prompt.
    3. Else the employee prompt is appended at the end after 【待检查文本】.
- str.format() is deliberately not used, so braces {} inside the Prompt or the
  employee text will not break anything.
- The run log prints which of the three rules above was applied.

Result columns / resuming
-------------------------
- 结果列1 and 结果列2 must be two different columns and must differ from 文字列.
- With "跳过已有结果的单元格" checked:
    * a row whose result columns are ALL filled is skipped;
    * a row with only ONE filled column only runs the missing column again
      (useful after an interrupted run).
- This tool always reads the INPUT file, so to resume an interrupted run, pick
  the previously generated *_AI_result.xlsx as the input file.

How to start on Windows
-----------------------
1. Keep START_TOOL.cmd and company_ai_excel_tool.py in the same folder.
2. Double-click START_TOOL.cmd.
3. Select the input Excel file.
4. Paste the complete Authorization value (Bearer ...).
5. Choose the worksheet / input column / result column 1 / result column 2 /
   start row.
6. Leave Prompt blank to use DEFAULT_PROMPT, or enter a custom Prompt.
7. Click Start.
8. Keep the output Excel file closed while the tool is running.

Dependencies
------------
Python 3, requests, openpyxl, tkinter (normally included with standard Python).

START_TOOL.cmd now checks these for you:
1. It looks for a usable interpreter ("py -3", then "python").
2. If it cannot import openpyxl / requests, it runs
   "<python> -m pip install openpyxl requests" automatically.
3. If the tool exits with an error, the console window STAYS OPEN and shows the
   reason, instead of closing instantly.
Keep START_TOOL.cmd saved as plain ASCII (no BOM) so the console output does not
turn into garbled text.

Troubleshooting
---------------
- Clicking START_TOOL.cmd and nothing happens:
  the window used to close before you could read the error. It now pauses on
  failure. The most common cause is a missing package - START_TOOL.cmd will try
  to install it for you.
- "No module named 'openpyxl'" / "No module named 'requests'":
  run the install manually with the SAME interpreter the launcher picks:
      py -3 -m pip install openpyxl requests
  Note that a different Python than "py -3" is a different environment - every
  interpreter needs its own copy of the packages.
- "Python 3 was not found":
  install Python 3 and make sure "py" or "python" works in a Command Prompt.
- Manual launch, if you want to see the console output yourself:
      cd /d <folder>
      py -3 company_ai_excel_tool.py

Security / internal-use note
----------------------------
The token is entered at runtime and is not intentionally written to the Excel,
config files, or logs. Confirm that automated and concurrent access to the
internal AI API is permitted by your company's policies.
