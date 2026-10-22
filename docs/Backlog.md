---
tags: [backlog]
---
# Backlog

Back to [[Home]].

## Done 2026-10-22
- [x] QA polish on the live workbook via `office-scripts/Polish (one-off).ts` (K5 alignment, K6 border, Instructions wording).
- [x] Mirrored the live Instructions wording into the local `.xlsx` (online wording, snapshot section added) and `.xlsm` (desktop wording).

## Done 2026-10-22 (later)
- [x] Start/End date hint + calendar icon, date-only validation; New Project start date optional (blank = today).
- [x] Example text under every form field (online `@param` hints; desktop InputBox prompts in `vba/*.bas`).
- [x] Sheet protection on the live file and the local `.xlsx`; scripts unlock/relock.

## Open
- [ ] Desktop `.xlsm`: import the updated `vba/ProjectTools.bas` / `GanttTools.bas` (prompt examples) and add
      protection with `UserInterfaceOnly:=True` in Workbook_Open so macros still run. Needs "Trust access to the
      VBA project object model" switched on by Hayden, or a manual Alt+F11 import.

## Next
- [ ] Fill in Atlantic Packaging's Client Email (Projects K6).
- [ ] Freeze panes on project tabs so task names and buttons stay visible when scrolling the chart.
- [ ] Move the live workbook from Hayden's OneDrive to a SharePoint/Teams site (then re-point the flow – see [[Email Snapshot Flow]]).
- [ ] Consider a run-only/service account for the flow so emails don't come from Hayden.

## Ideas
- [ ] TigerPaw service-order import (Power Automate HTTP → Office Script that creates a project). Needs a sample service order and the field mapping.
- [ ] Customer workbook per project kept in sync (alternative to snapshots).
- [ ] Delete Rows equivalent online (Office Scripts can't confirm; could use a parameter checkbox).
