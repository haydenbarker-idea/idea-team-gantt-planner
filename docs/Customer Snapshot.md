---
tags: [snapshot]
---
# Customer Snapshot

Back to [[Home]]. Same rules in two places: `vba/SnapshotTools.bas` (desktop) and
`office-scripts/Prepare Snapshot.ts` (online, run by [[Email Snapshot Flow]]).

Always works on a copy, never the live workbook:
1. Only the one project tab (desktop: `Worksheet.Copy` to a new workbook; online: delete the other sheets in a temp copy).
2. Remove buttons, comments, the "All projects" link (I1), invoice labels (I3:ED3), Display Week (D4:E4); reset row-5 day fills.
3. E4 = 1, recalculate, then values only. D3 reads "As of:" with the export date.
4. Crop the chart to whole weeks up to the last end date (max 18 weeks); delete rows below the last task.
5. Dates "ddd mmm d, yyyy", E:F narrower.
6. Page setup: landscape, fit 1 page wide, print titles $4:$6, Letter (Tabloid if wider than ~1100 pt), centred,
   footer "Idea Networks · Page &P of &N · Schedule as of <date>".
7. Files: `<Project> - Schedule <yyyy-mm-dd>.pdf` and `.xlsx`.

Checked by reading the PDFs for Atlantic Packaging: one clean page; online and desktop output match.

Clients get files only. Sharing a single sheet or embedding was rejected – see [[Decisions]].
