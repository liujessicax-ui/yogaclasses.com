---
name: log-yoga-bugs
description: Log user-identified bugs for the yoga site (yogawithjessica.com booking system) into bug-tracker.xlsx. Use ONLY when Jessica explicitly says to log the bugs (e.g. "log the bugs to yoga site", "log this bug to the yoga site", "add that to the bug tracker"). For each bug she identified, record her original description, an enhanced/corrected description, the root cause, and the resolution.
---

# Log Yoga Site Bugs

Records bugs **Jessica has identified** into the bug tracker at the repo root:
`bug-tracker.xlsx` (sheet "Bug Tracker").

## When to use
Only when the user explicitly asks to log bug(s) to the yoga site — e.g. "log the
bugs to yoga site", "log that bug", "add these to the bug tracker". Do **not** log
anything proactively or as part of fixing a bug; logging only happens on command.

## Hard rules
- **Only log bugs the USER identified.** Never log a bug you discovered on your own
  (e.g. an issue you noticed while reading code). If you found something worth
  tracking, mention it to the user and let *them* decide to log it.
- **Never edit or delete existing rows.** This skill only appends.
- **Never invent details.** Root cause and resolution must come from the actual
  investigation/fix in this conversation (or a fresh investigation — see below).
  If you don't know a field, say so rather than guessing.
- One row per distinct bug. If the user reported several in one breath (as with the
  three booking bugs), log each as its own row.

## Which bugs to log
1. Default to the bug(s) the user is referring to **right now** — usually the ones
   just discussed in this conversation.
2. If it's ambiguous which bug(s) they mean, **ask** before writing anything.
3. Before logging, read the current rows so you don't duplicate an existing entry
   (compare by the original description / root cause):
   ```bash
   python3 -c "from openpyxl import load_workbook; ws=load_workbook('bug-tracker.xlsx')['Bug Tracker']; [print(r[0],'|',r[3]) for r in ws.iter_rows(min_row=2,values_only=True) if r[0]]"
   ```

## What each field must contain
For every bug, produce these four core fields:

- **Original Description (Jessica's words)** — Capture *how she described it*, faithfully.
  Quote or closely paraphrase her own report, including any misunderstanding she had
  about what was wrong (that contrast is the point of the next field). Do not "fix" her
  wording or silently correct mistaken assumptions here.
- **Enhanced Description (refined understanding)** — Your corrected, precise restatement
  of the actual problem. If her original framing was off (e.g. she called it a
  "night-before email" when it was the morning-of reminder, or assumed a "33-minute"
  signup edge when the real trigger was a window-coordination gap), explain the
  correction plainly here.
- **Root Cause** — The actual technical cause, from investigation. Be specific
  (file/function/mechanism), e.g. "Script pasted into the Apps Script editor re-decoded
  UTF-8 em-dash bytes as Mac Roman."
- **Resolution / Fix Applied** — Concretely what was done to correct it, or operational
  steps if there was no code change. If the bug is still open, leave this empty and set
  status to Open / In Progress.

Also fill where known:
- **Status** — Open | In Progress | Fixed | Won't Fix | Not a Bug. (If omitted, the
  helper defaults to "Fixed" when a resolution is present, else "Open".)
- **Reference** — commit hash, PR link, and/or files touched.

`Bug ID` and `Date Logged` are assigned automatically — do not set them.

## If a bug hasn't been analyzed yet
If the user points at a bug that hasn't been investigated in this conversation, do the
analysis FIRST: read the relevant code, reproduce/trace the issue, and determine the
real root cause before writing the row. The enhanced description, root cause, and
resolution must reflect genuine findings — not a restatement of the symptom.

## How to write the rows
1. Build a JSON file describing the bug(s). Write it to the scratchpad, not the repo.
   It is a list of objects (one per bug):
   ```json
   [
     {
       "original": "Jessica's own description of the bug",
       "enhanced": "Refined, corrected understanding of the real problem",
       "root_cause": "The actual technical cause",
       "resolution": "What was done to fix it (empty string if still open)",
       "status": "Fixed",
       "reference": "commit 1a2b3c / PR #4 / google-apps-script.js"
     }
   ]
   ```
2. Run the helper (resolves `bug-tracker.xlsx` at the repo root automatically):
   ```bash
   python3 .claude/skills/log-yoga-bugs/append_bug.py /path/to/bugs.json
   ```
   It auto-assigns the next `BUG-NNN` id and today's date, preserves formatting, and
   prints what it logged.
3. If openpyxl is missing, install it: `pip3 install openpyxl --quiet`.
4. Report back to the user: the Bug ID(s) assigned and a one-line summary of each, so
   she can confirm the entries are accurate.

## Spreadsheet shape (for reference)
Columns A–I: Bug ID · Date Logged · Reported By · Original Description (Jessica's words) ·
Enhanced Description · Root Cause · Resolution / Fix Applied · Status · Reference.
Header row is frozen; Status has a dropdown.
