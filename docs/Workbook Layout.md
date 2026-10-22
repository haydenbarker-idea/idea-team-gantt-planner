---
tags: [workbook]
---
# Workbook Layout

Back to [[Home]].

## Tabs
| Tab | Visible | Purpose |
|---|---|---|
| Projects | yes | One row per project (list starts row 6). |
| Instructions | yes | How-to, worded per version (web vs desktop). |
| project tabs | yes | One Gantt per project (e.g. "Atlantic Packaging"). |
| Template | hidden | Copied by New Project. Has the row-1 buttons so new projects inherit them. |
| About | hidden | Original Vertex42 credits. |
| Snapshot Requests | hidden, online only | `SnapshotRequests` table used by [[Email Snapshot Flow]]. |

## Projects tab columns (headers row 5, data from row 6)
- B Project – HYPERLINK to the tab (uses J), text = tab's B1
- C Client · D Project manager (= tab B3)
- E Start = MIN(tab E7:E1000) · F Finish = MAX(tab F7:F1000)
- G Progress = SUMPRODUCT(D, H) / SUM(H) over the tab (weighted by task days)
- H Status: Archived / Not scheduled / Complete / Not started / Overdue / In progress
- I Archived ("Yes"/blank) · J Tab name (typed) · K Client Email (typed)

Ranges start at row 7 (a blank spacer) so rows inserted at the top of a task list stay inside them.
Row 1: title + buttons over E–G. B3: "How to use this workbook ►" link to Instructions.

## Project tab
- B1 project name, B2 company ("Idea Networks"), B3 project manager.
- E2 project start, E3 `=TODAY()`, E4 Display Week (scrolls chart; 1 = start week).
- Row 1: "◄ All projects" link in merged I1:N1, then buttons.
- Row 3 chart area: internal invoice labels; row 5 day numbers (Atlantic has hand-coloured orange invoice days; removed from the Template).
- Row 6 headers; H = DAYS (hidden, width 0); tasks from row 8 to the grey "Insert new rows ABOVE this one" row.
- Section header = bold B with empty E; task = has a name or start date.
- Chart columns I..ED (126 days). Bars are conditional formatting using sheet-scoped names
  `task_start` ($E), `task_end` ($F), `task_progress` ($D), `today` ($E$3).
- Section colour pairs (header / task): #B8CCE4/#DCE6F1, #D8E4BC/#EBF1DE, #E6B8B7/#F2DCDB,
  #CCC0DA/#E4DFEC, #B7DEE8/#DAEEF3. New sections pick one different from both neighbours.

## Protection and date entry (live file and local `.xlsx`; not yet the `.xlsm`)
No password: it guards against accidents. Set up by `office-scripts/Dates and Protection (one-off).ts`.
- Project tabs + Template: everything locked except B1, B3, E2, E4, I3:ED3 and whole rows 7 down to
  just above the marker (whole rows, so they can still be deleted). Insert/delete rows and formatting
  allowed; inserting/deleting columns, sorting, filtering not. Objects editable (buttons need it).
- Projects: everything locked except Client (C), Tab (J), Client Email (K). No row insert/delete.
- Instructions: fully locked. Snapshot Requests: unprotected (the flow writes to it).
- Start/End (E7:F marker) take dates only (2000–2099), with an input hint "Double-click for a calendar";
  E6/F6 headers have a 📅. Inserted rows inherit the rule (the range includes the marker row).
- Scripts lift protection on the sheet they change and put the same options back (`unprotected()` helper).

See [[Customer Snapshot]] for how a tab is reduced for clients.
