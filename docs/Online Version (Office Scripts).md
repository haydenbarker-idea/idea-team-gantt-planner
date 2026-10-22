---
tags: [office-scripts, online]
---
# Online Version (Office Scripts)

Back to [[Home]]. Live file: [[Deployment and Sharing]]. Code: `office-scripts/`.
Saved copies: Hayden's OneDrive > Documents > Office Scripts (`.osts` = JSON with `body`, `parameterInfo`).

| Script | Where | Notes |
|---|---|---|
| New Project | Projects tab | projectName*, startDate* (2026-11-02 or 11/2/2026), client?, projectManager?, clientEmail? |
| Archive Project | Projects tab | Select a project row; toggles Archived and hides/unhides the tab. |
| Add Section / Add Task | project tabs + Template | Same rules as the VBA version. |
| Email Snapshot to Client | project tabs + Template | Queues a request for [[Email Snapshot Flow]]. |
| Prepare Snapshot | flow only | Refuses unless the workbook name starts `snapshot-tmp-`. |
| Setup Snapshot (one-off), Polish (one-off) | run once on the live file | Kept for reference; their `.osts` copies are deleted after running. |

## Working with scripts in Excel for the web
- Create: Automate > New Script > Create in Code Editor → paste → rename (pen icon) → save. Check the name stuck (the title sometimes stays "Script").
- Share: open from the Automate gallery → **Associate with workbook** → **Add button to worksheet** (lands at the selected cell; right-click a button to select it without running).
- Buttons on the Template are inherited by new projects (tested).
- Position buttons precisely with a script (`shape.setLeft/Top/Width/Height`), not by dragging.
- No confirm dialogs and no current-user API. Requester email: add a comment, read `getAuthorEmail()`, delete it.
- Body-only changes can be made by editing the `.osts` `body` in the synced folder; parameter changes must go through the editor.

See [[Lessons Learned]].
