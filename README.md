# Project Schedules – Gantt workbook

| File | Use it for |
|---|---|
| `Project Schedules.xlsm` | The team workbook with buttons, for desktop Excel (VBA macros). |
| `Project Schedules.xlsx` | The same workbook without macros, for SharePoint / Excel for the web with the Office Scripts in `office-scripts/` (the scripts and their buttons are added in Excel for the web, see below). |
| `Atlantic Schedule.xlsm` / `.xlsx` | The original single-project schedule (now superseded; Atlantic is the first project in Project Schedules). |

The live team copy is in OneDrive > My files > Project Schedules (test) > Project Schedules.xlsx,
shared with Vinny Singh (edit).

## How the workbook is laid out

- **Projects** tab: one row per project with client, project manager, start, finish,
  progress (weighted by task length), status (Not started / In progress / Overdue /
  Complete / Archived) and Client Email. Click a project name to open its tab. Only Client,
  Archived, Tab and Client Email are typed in; the rest is read from the project's tab.
- **One tab per project**, each a Gantt chart. "◄ All projects" (top of the chart) goes back
  to the Projects tab.
- **Template** (hidden): what new projects are copied from. Unhide it (right-click a tab >
  Unhide) to change the starting sections, colours or layout for future projects.
- **Instructions**: the how-to, worded for that file's version (web buttons in the `.xlsx`,
  macro buttons in the `.xlsm`). "How to use this workbook ►" on the Projects tab opens it.
- **About** (hidden): the original Vertex42 template credits, kept but out of the way.
- **Snapshot Requests** (hidden, online only): the queue the Email Snapshot button writes to.

If you rename a project tab, put the new name in its **Tab** cell on the Projects tab so the
link, Archive and the snapshot email keep working.

## Buttons

On the **Projects** tab:

| Button | What it does |
|---|---|
| **New Project** | Asks for project name, start date and (optionally) client, project manager and client email; creates the tab from the Template and adds it to the list. |
| **Archive / Restore** | For the selected project in the list (or the tab you're on): marks it Archived and hides its tab. On an archived project it unhides it again. |

On each **project tab** (row 1, above the chart):

| Button | What it does |
|---|---|
| **Add Section** | Asks for a name, then adds a section header with one "New Task" under it, after the section you have selected (or at the end). New sections get a colour that differs from their neighbours. |
| **Add Task** | Asks for a name, then adds a task below the selected task, or at the end of the section if the header is selected. Dates default to the day after the task above. |
| **Delete Rows** (desktop only) | Deletes the selected task rows; a selected section header deletes the whole section. Asks first. Online: right-click the rows > Delete. |
| **Export Snapshot** (desktop only) | Saves `<Project> - Schedule <date>.pdf` and `.xlsx` of just this project to a folder you pick. |
| **Email Snapshot to Client** | Desktop: opens an Outlook draft to the project's Client Email with both files attached. Online: emails both files to whoever clicked (within ~2 minutes) to forward to the client. |

A snapshot is a customer copy of one project tab: values only (frozen as of today), no
buttons, links, invoice markers or Display Week, the chart cropped to the project's
weeks, one page wide (11x17 for long projects), footer "Idea Networks · Page X of Y ·
Schedule as of <date>".

New rows get the section's colours, the DAYS formula, the progress data bar and the Gantt bars.
Button actions can't be undone with Ctrl+Z.

## VBA (desktop Excel) – `vba/`

`Project Schedules.xlsm` contains `GanttTools.bas`, `ProjectTools.bas` and `SnapshotTools.bas`.
To update the code after editing the `.bas` files: Alt+F11, remove the old module, File > Import File.

If Excel shows "Microsoft has blocked macros" for a copy from SharePoint or email:
right-click the file > Properties > tick Unblock, or open it from a trusted location.
The VBA buttons don't run in Excel for the web or Teams.

## Office Scripts (Excel for the web / Teams) – `office-scripts/`

For `Project Schedules.xlsx` on OneDrive / SharePoint. Scripts live in Hayden's
OneDrive > Documents > Office Scripts and are associated with the workbook, so anyone who
can edit the workbook can run them.

| Script | Button / use |
|---|---|
| `New Project.ts`, `Archive Project.ts` | Projects tab buttons. |
| `Add Section.ts`, `Add Task.ts` | Project tab buttons (also on the Template, so new projects get them). |
| `Email Snapshot to Client.ts` | Project tab button. Adds a row to the hidden SnapshotRequests table (Id, time, tab, requester, Status). The requester's email is read from a throwaway comment. |
| `Prepare Snapshot.ts` | Run only by the flow, on a temporary copy (refuses any file not named `snapshot-tmp-...`). Deletes the other tabs and builds the customer view. |
| `Setup Snapshot (one-off).ts` | Already run on the live workbook (Client Email header, Template markers, Instructions section, button alignment); kept for reference. |

Setting up a script: Automate > New Script > Create in Code Editor, paste, rename with the
pen icon, save; open it from the Automate gallery, turn on **Associate with workbook**, then
**Add button to worksheet** (the button appears at the selected cell). Make changes to the
shared workbook in Excel for the web: a save from desktop Excel briefly broke the buttons
online until the file finished syncing.

### Email snapshot flow (Power Automate)

Flow **"Project Schedules - email snapshot"** in Hayden's Power Automate (default
environment), scheduled every minute, using Hayden's Excel, OneDrive and Outlook connections:

1. List rows in SnapshotRequests where Status = Pending.
2. For each: skip it if `_snapshot-log/<Id>.done` exists (Excel reads can lag writes by a
   minute or two, so Status alone let requests be sent repeatedly); otherwise create that
   marker, mark the row Working, copy the workbook to `_snapshot-temp/snapshot-tmp-<guid>.xlsx`,
   run Prepare Snapshot on the copy, convert it to PDF, email the PDF and the copy to the
   requester, mark the row Sent (or Failed), and delete the copy (after 90 s, retried after 4 min
   because Excel Online keeps it open for a while).

The flow's file and table are the live workbook's IDs; the loop was generated with
`power-automate/make-loop.py` as a classic-designer clipboard item
(`power-automate/loop-live-workbook.json`) and pasted via + New step > My clipboard. If the workbook
moves, re-pick the file/table in "List rows" and regenerate the loop with the new IDs.
Turn the flow off (flow details > Enabled) to stop it.
