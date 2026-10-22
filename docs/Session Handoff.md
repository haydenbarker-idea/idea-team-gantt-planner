---
tags: [handoff]
updated: 2026-10-22
---
# Session Handoff (2026-10-22)

Start here in a new session, then [[Home]] and [[Lessons Learned]].

## State
- Repo: `C:\Dev\gantt-chart-varennes` → GitHub `haydenbarker-idea/idea-team-gantt-planner` (main, pushed).
  Push uses deploy key `~/.ssh/idea-team-gantt-planner_deploy` via ssh alias `github-gantt-planner` (already configured).
- Live workbook (source of truth): Hayden's OneDrive > Project Schedules (test) > `Project Schedules.xlsx`.
  Shared with Vinny Singh (edit). Tabs: Projects, Instructions, Atlantic Packaging (+ hidden Template, About, Snapshot Requests).
- Online buttons working and QA-passed: New Project, Archive Project, Add Section, Add Task, Email Snapshot to Client.
- Live sheets are protected (no password) and Start/End have a calendar hint; see [[Workbook Layout]]. Projects: Atlantic
  Packaging (sample) and Costco Varennes 573 Access Control (real, made by Hayden).
- Flow "Project Schedules - email snapshot" ON, every minute, owned by Hayden ([[Email Snapshot Flow]]).
- Desktop `Project Schedules.xlsm` in the repo: all buttons incl. Export Snapshot / Email Snapshot (Outlook draft).
- Local `.xlsx`/`.xlsm` are reference copies; their Instructions match the live wording (desktop wording in `.xlsm`).

## Open items ([[Backlog]])
1. Fill Atlantic Packaging's Client Email (Projects!K6).
2. Desktop `.xlsm`: VBA prompt examples + protection not imported yet ([[Backlog]]).
3. Optional: freeze panes on project tabs; move live file to SharePoint/Teams (then re-point flow);
   service account for the flow; TigerPaw service-order import (needs a sample order).

## How to work safely
- Change the live workbook only in Excel for the web, via one-off Office Scripts (pattern:
  `office-scripts/Setup Snapshot (one-off).ts`, `Polish (one-off).ts`), then run a QA sub-agent
  (read-only, never click Email Snapshot or Archive), then delete the one-off `.osts` from
  OneDrive > Documents > Office Scripts.
- Desktop tests via Excel COM on scratch copies; set `ActivePrinter` to "Microsoft Print to PDF on Ne02:" first.
- Chrome automation: keep the Chrome window in front (screenshots time out otherwise).
- Flow edits: classic designer + `power-automate/make-loop.py` clipboard JSON.
