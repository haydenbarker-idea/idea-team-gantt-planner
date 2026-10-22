---
tags: [decisions]
---
# Decisions

Back to [[Home]].

| Decision | Why |
|---|---|
| One tab per project + Projects overview + hidden Template | Keeps the original Gantt; copying a sheet carries its formatting, names and buttons. A single table + filtered Gantt would be a rebuild; revisit past ~30–40 live projects. |
| Two versions: `.xlsm` (VBA) and `.xlsx` (Office Scripts) | VBA doesn't run in Excel for the web/Teams; the team uses the shared online file. |
| Edit the live file only in Excel for the web | Desktop saves disturbed script associations ([[Lessons Learned]]). |
| Customer gets a snapshot file, not access | Sharing is per file; hiding tabs or embedding doesn't protect other projects. |
| Online snapshot = queue + scheduled flow, emailed to the requester to forward | Office Scripts can't create files/emails; a flow can only draft in its owner's mailbox, so the user chose "emailed to them to forward". |
| Marker files for de-duplication | Excel reads are stale for 1–2 min. |
| Prepare Snapshot refuses non-`snapshot-tmp-` files | It deletes tabs; guard against running it on the real workbook. |
| Script-positioned buttons | Dragging is imprecise; QA flagged misalignment. |
| TigerPaw service-order import deferred | User chose to do the multi-project workbook first ("leave step 2"). |
