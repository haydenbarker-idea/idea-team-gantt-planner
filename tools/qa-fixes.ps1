# QA fixes: clickable "All projects" links, Instructions wording per version, hide About.
# Usage: qa-fixes.ps1 <workbook path> <web|desktop>
param([string]$Path, [string]$Variant)

$blue = 5 + 256 * 99 + 65536 * 193
$gray = 89 + 256 * 89 + 65536 * 89
$navy = 68 + 256 * 84 + 65536 * 106
$la = [string][char]0x25C4; $ra = [string][char]0x25BA; $bullet = [string][char]0x2022
$fAll = [string]('=HYPERLINK("#''Projects''!A1","' + $la + ' All projects")')

if ($Variant -eq "web") {
  $btn = @{ new = "Run New Project"; archive = "Run Archive Project"; restore = "Run Archive Project (again)"; task = "Run Add Task"; section = "Run Add Section" }
  $lines = @(
    "t|How to use this workbook",
    "s|One Gantt chart per project, plus a Projects overview. Idea Networks.",
    "x|",
    "h|Finding your way around",
    "b|Projects tab: one line per project with client, project manager, start, finish, progress and status. Click a project name to open its tab.",
    "b|One tab per project: the Gantt chart. The '$la All projects' link at the top of the chart takes you back to the Projects tab.",
    "b|Template (hidden tab): new projects are copied from it. To change what new projects start with, right-click any tab > Unhide > Template, edit it, then hide it again.",
    "x|",
    "h|Start a new project",
    "n|1.  Go to the Projects tab and click Run New Project.",
    "n|2.  Enter the project name and start date (for example 2026-11-02); client and project manager are optional. Click Run.",
    "n|3.  The new tab opens with one 'Planning' section and one task. Rename them and add your own sections and tasks.",
    "x|",
    "h|Build the schedule",
    "b|Run Add Task: select a task first; the new task goes right below it. Select a section header instead and it goes at the end of that section.",
    "b|Run Add Section: select anything in a section; the new section goes after it.",
    "b|To delete tasks, select their rows, right-click > Delete. Delete a section the same way (its header row plus its tasks).",
    "b|Type task names, Assigned To, Progress (%), Start and End straight into the cells. The bars, DAYS and the Projects tab update by themselves.",
    "b|Display Week (cell E4) moves the chart: 1 is the week the project starts, 2 the week after, and so on.",
    "x|",
    "h|Finish or put away a project",
    "b|On the Projects tab, click the project's row, then Run Archive Project. Its tab is hidden and its status shows Archived.",
    "b|To bring it back, click the archived row and Run Archive Project again.",
    "x|",
    "h|Good to know",
    "b|The buttons run Office Scripts. Each one asks for its inputs in a small panel; click Run to go ahead or Cancel to stop. They are also on the Automate tab of the ribbon.",
    "b|Script actions can't always be undone with Ctrl+Z, so take care with Archive and deleting.",
    "b|Keep all tasks above the grey 'Insert new rows ABOVE this one' row.",
    "b|If you rename a project tab, type the new name in its Tab cell on the Projects tab so the link and Archive keep working.",
    "b|Progress on the Projects tab weights each task by its length in days. Status turns Overdue when today is past the last end date and progress is under 100%.",
    "b|Everyone should work in this one shared file (from the link, in the browser or Teams), not a downloaded copy."
  )
} else {
  $lines = @(
    "t|How to use this workbook",
    "s|One Gantt chart per project, plus a Projects overview. Idea Networks. Desktop Excel version (macros).",
    "x|",
    "h|Finding your way around",
    "b|Projects tab: one line per project with client, project manager, start, finish, progress and status. Click a project name to open its tab.",
    "b|One tab per project: the Gantt chart. The '$la All projects' link at the top of the chart takes you back to the Projects tab.",
    "b|Template (hidden tab): new projects are copied from it. To change what new projects start with, right-click any tab > Unhide > Template, edit it, then hide it again.",
    "x|",
    "h|Start a new project",
    "n|1.  Go to the Projects tab and click + New Project.",
    "n|2.  Enter the project name, client, project manager and start date (for example 2026-11-02).",
    "n|3.  The new tab opens with one 'Planning' section and one task. Rename them and add your own sections and tasks.",
    "x|",
    "h|Build the schedule",
    "b|+ Add Task: select a task and click; the new task goes right below it. Select a section header instead and it goes at the end of that section.",
    "b|+ Add Section: select anything in a section and click; the new section goes after it. With nothing in the task list selected, it goes at the end.",
    "b|Delete Rows: select the tasks to delete (Ctrl+click for several) and click. Selecting a section header deletes the whole section. You'll be asked to confirm.",
    "b|Type task names, Assigned To, Progress (%), Start and End straight into the cells. The bars, DAYS and the Projects tab update by themselves.",
    "b|Display Week (cell E4) moves the chart: 1 is the week the project starts, 2 the week after, and so on.",
    "x|",
    "h|Finish or put away a project",
    "b|On the Projects tab, select the project's row and click Archive / Restore. Its tab is hidden and its status shows Archived.",
    "b|To bring it back, select the archived row and click Archive / Restore again.",
    "x|",
    "h|Send the customer a snapshot",
    "b|Email Snapshot to Client (on the project's tab): makes a PDF and an Excel copy of just that project and opens an Outlook draft to the client with both attached. Check it and press Send.",
    "b|Export Snapshot: saves the same PDF and Excel copy to a folder you choose.",
    "b|The snapshot is a frozen copy as of today: no other projects, no formulas, no invoice markers or internal controls. The client's email comes from the Client Email column on the Projects tab.",
    "x|",
    "h|Good to know",
    "b|Button actions can't be undone with Ctrl+Z. Save before big changes.",
    "b|Keep all tasks above the grey 'Insert new rows ABOVE this one' row.",
    "b|If you rename a project tab, type the new name in its Tab cell on the Projects tab so the link and Archive / Restore keep working.",
    "b|Progress on the Projects tab weights each task by its length in days. Status turns Overdue when today is past the last end date and progress is under 100%.",
    "x|",
    "h|Opening the macros",
    "b|If a yellow bar says 'Enable Content', click it.",
    "b|If a red bar says 'Microsoft has blocked macros', close the file, right-click it > Properties > tick Unblock > OK, then open it again.",
    "b|The macro buttons only work in desktop Excel. For Excel for the web / Teams, use the online version of this workbook."
  )
}

$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open([string]$Path)
  $out = @()

  # Project tabs: make the whole "All projects" text clickable (I1:N1 merged, before the buttons).
  foreach ($s in @($wb.Worksheets)) {
    if ([string]$s.Range("B6").Value2 -ne "TASK") { continue }
    $area = $s.Range("I1").MergeArea
    $j = $s.Range("J1").MergeArea
    if ($j.Cells.Count -gt 1) { [void]$j.UnMerge() }
    if ($area.Cells.Count -gt 1) { [void]$area.UnMerge() }
    [void]$s.Range("I1:N1").Merge()
    $c = $s.Range("I1"); $c.Formula = $fAll; $c.HorizontalAlignment = -4131; $c.VerticalAlignment = -4108
    $c.Font.Size = 10; $c.Font.Underline = 2; $c.Font.Color = $blue
    $out += "$($s.Name): link merged I1:N1"
  }

  # Instructions: rewrite text for this version; link in C3 (the wide column).
  $ins = $wb.Worksheets.Item("Instructions")
  [void]$ins.Range("A1:D200").Clear()
  $r = 1
  foreach ($line in $lines) {
    $kind, $text = $line -split '\|', 2
    $text = [string]$text
    switch ($kind) {
      "t" { $c = $ins.Range("B$r"); $c.Value2 = $text; $c.Font.Size = 20; $c.Font.Bold = $true; $ins.Rows($r).RowHeight = 30 }
      "s" { $c = $ins.Range("B$r"); $c.Value2 = $text; $c.Font.Color = $gray; $ins.Rows($r).RowHeight = 15 }
      "h" { $c = $ins.Range("B$r"); $c.Value2 = $text; $c.Font.Size = 13; $c.Font.Bold = $true; $c.Font.Color = $navy; $ins.Rows($r).RowHeight = 22 }
      "b" { $ins.Range("B$r").Value2 = $bullet; $ins.Range("B$r").HorizontalAlignment = -4152; $c = $ins.Range("C$r"); $c.Value2 = $text; $c.WrapText = $true }
      "n" { $c = $ins.Range("C$r"); $c.Value2 = $text; $c.WrapText = $true }
      "x" { $ins.Rows($r).RowHeight = 8 }
    }
    $ins.Range("B$($r):C$r").VerticalAlignment = -4160
    $r++
  }
  foreach ($row in 1..$r) {
    $k = ($lines[$row - 1] -split '\|', 2)[0]
    if ($k -eq "b" -or $k -eq "n") { [void]$ins.Rows($row).AutoFit() }
  }
  # Row 3 is a spacer; put the back link there in the wide column C.
  $ins.Rows(3).RowHeight = 18
  $c = $ins.Range("C3"); $c.Formula = $fAll; $c.Font.Size = 10; $c.Font.Underline = 2; $c.Font.Color = $blue; $c.VerticalAlignment = -4108
  $ins.Range("B1:C$r").Font.Name = [string]$wb.Worksheets.Item("Projects").Range("B6").Font.Name
  $out += "Instructions rewritten ($Variant), $($r - 1) lines, link in C3"

  # Old Vertex42 credits: keep in the file but out of the way.
  foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq "About") { $s.Visible = 0; $out += "About hidden" } }

  $wb.Worksheets.Item("Projects").Activate()
  $wb.Save(); $wb.Close($false)
  $out; "saved $Path"
} catch { "ERROR: $($_.Exception.Message) at $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" }
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
