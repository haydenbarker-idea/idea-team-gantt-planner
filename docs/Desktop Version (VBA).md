---
tags: [vba, desktop]
---
# Desktop Version (VBA)

Back to [[Home]]. File: `Project Schedules.xlsm`. Code: `vba/`.

| Module | Buttons | Work routines (callable via Application.Run) |
|---|---|---|
| `GanttTools.bas` | AddSection, AddTask, DeleteRows | InsertSection(name), InsertTask(name), DeleteRows(askFirst) |
| `ProjectTools.bas` | NewProject, ArchiveProject | CreateProject(name, client, pm, start, [clientEmail]), SetArchived(row, bool) |
| `SnapshotTools.bas` | ExportSnapshot, EmailSnapshot | SaveSnapshot(ws, folder), SaveSnapshotFor(tab, folder), DraftSnapshotEmailFor(tab, show), SnapshotErrorText() |

- Works on any tab with "TASK" in B6; falls back to Template formatting if a tab has no sections left.
- Buttons: Form Controls, free-floating, not printed, 110×22 pt in row 1 from I1.Left+90, 8 pt gaps.
- Email Snapshot opens an Outlook draft (with the user's signature) to the Client Email (Projects K); it never sends.
- Snapshot errors are reported with the failing step (e.g. "saving the PDF: ...").

Installing/updating code: Alt+F11 → remove module → File > Import File. Doing it by automation
needs "Trust access to the VBA project object model" (the user turns it on and off).

See [[Customer Snapshot]], [[Testing and QA Log]], [[Lessons Learned]].
