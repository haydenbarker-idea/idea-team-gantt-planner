---
tags: [testing]
---
# Testing and QA Log

Back to [[Home]]. All 2026-10-22.

## Desktop (COM automation on scratch copies)
- Add Task / Add Section / Delete Rows (incl. multi-select, whole section): rows, colours, DAYS, bars, CF rule count stable.
- New Project (duplicate names → "(2)", invalid chars stripped, 31-char limit), Archive/Restore, Template fallback when a tab is emptied.
- Projects formulas update when rows are added at the top of a task list (ranges start row 7).
- Snapshot: PDF and xlsx checked (one page, no formulas/shapes, values frozen); Outlook draft with both attachments (test draft deleted).
- Tests set the hidden Excel's printer to "Microsoft Print to PDF" to avoid driver pop-ups.

## Online (Excel for the web via Chrome)
- All scripts run from the editor and from sheet buttons; parameter dialogs; Template buttons inherited by new projects.
- QA agent passes (independent sub-agent, read-only, cancels dialogs):
  1. Layout + links: found links only clickable on the arrow → merged link cells; Instructions wording → rewritten per version; About hidden.
  2. Found buttons failing ("script was unshared") right after a desktop save → transient, recovered after sync.
  3. Snapshot rollout: all checks pass; minor polish items → [[Backlog]].
- 2026-10-22 dates + protection (live): Start/End hint and double-click calendar on protected tabs; locked cell (E3 Today)
  refused; E2/B1 editable; Run Add Task on protected Atlantic tab (needed objects editable) then manual row delete
  (needed whole rows unlocked); New Project form shows examples and optional start date (cancelled, not run).
  Local: inserted rows inherit the date rule (COM on a scratch copy).
- Email Snapshot flow: first live test sent 6 duplicates (stale reads) → flow turned off, marker files added → retest sent exactly 1 (test copy), then exactly 1 on the live workbook (10:27).
