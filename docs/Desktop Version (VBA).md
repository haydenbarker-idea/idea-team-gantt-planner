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
| `ProtectTools.bas` | – | ProtectAll(), ProtectSchedule(ws) |
| `ThisWorkbook.cls` | – | Workbook_Open → ProtectAll (code pasted into ThisWorkbook, not imported) |

- Sheets are protected like the online file ([[Workbook Layout]]) with UserInterfaceOnly, which Excel doesn't save:
  Workbook_Open and every button's work routine call ProtectAll first. New tabs get ProtectSchedule; snapshot copies are unprotected.
- Works on any tab with "TASK" in B6; falls back to Template formatting if a tab has no sections left.
- Buttons: Form Controls, free-floating, not printed, 110×22 pt in row 1 from I1.Left+90, 8 pt gaps.
- Email Snapshot opens an Outlook draft (with the user's signature) to the Client Email (Projects K); it never sends.
- Snapshot errors are reported with the failing step (e.g. "saving the PDF: ...").

Installing/updating code: Alt+F11 → remove module → File > Import File. Doing it by automation
needs "Trust access to the VBA project object model" (the user turns it on and off). The 2026-10-22 install
script removed and re-imported all four modules and rewrote ThisWorkbook in one pass.

See [[Customer Snapshot]], [[Testing and QA Log]], [[Lessons Learned]].
