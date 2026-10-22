---
tags: [gotchas]
---
# Lessons Learned

Back to [[Home]].

## Excel / Office Scripts
- **Don't save the live online workbook from desktop Excel.** It briefly broke the script buttons ("script was unshared", blank dialogs) until sync settled.
- **Excel connector reads lag writes by a minute or two.** Status columns can't de-duplicate a 1-minute flow → use marker files in OneDrive ([[Email Snapshot Flow]]).
- **Excel Online keeps a file open after Run script**; deleting it straight away gives BadRequest → wait, retry later.
- **Merged cells swallow writes**: E2:F2 and E3:F3 are merged; writing F3 silently does nothing.
- **Narrow-column hyperlinks** are only clickable on the part inside the cell → merge the link cell wide enough.
- **Automate pane quirks**: script title can stay "Script" after rename; Recent list only shows recent scripts; Template must be unhidden to add buttons.
- **Office Scripts has no current-user API** → comment author trick.
- **OneDrive "Convert file"** honours the sheet's print area and fit-to-page settings.
- **`Range.find` returns nothing on hidden tabs** (the Template) → scan `getValues()` instead.
- **Script buttons do nothing on a protected sheet** unless protection allows editing objects.
- **A protected sheet won't delete a row that has any locked cell** (even far right of the chart) → unlock whole task rows.
- **Merged cells** (E2:F2): desktop COM can't set Locked on part of a merge → use `MergeArea`; Office Scripts handles it.
- **Chart looked blank right after opening** a tab in the browser: the bars render a moment later (seen on Costco, 2026-10-22).

## Automation from this PC
- PowerShell → Excel COM: string *variables* sometimes fail with "Specified cast is not valid" → cast with `[string]`; `Cells(r,c)` needs `.Item(r,c)`.
- Page setup in a hidden Excel triggers the Brother printer driver pop-ups on the user's screen → set `ActivePrinter` to "Microsoft Print to PDF on Ne02:" in tests; `Application.PrintCommunication` can throw → wrapped.
- `Application.Run` can't pass a Worksheet object from PowerShell → add name-based wrappers.
- Chrome screenshots time out when the Chrome window isn't in front; DOM tools (find/get_page_text) still work.
- Power Automate classic designer "Copy to my clipboard" JSON can be generated and pasted (much faster than building actions by hand); the new designer's clipboard format is internal UI state (not practical).
- Run long waits as background jobs; foreground sleeps are blocked.
