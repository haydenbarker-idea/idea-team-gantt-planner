Attribute VB_Name = "GanttTools"
Option Explicit

' Buttons for the project tabs: Add Section, Add Task, Delete Rows. They work on
' whichever project tab is active (any sheet with "TASK" in B6).
' A section header row is bold in column B with no start date (column E);
' a task row has a task name or a start date. Everything sits between
' FIRST_ROW and the "Insert new rows ABOVE this one" marker row.

Private Const TEMPLATE_SHEET As String = "Template"
Private Const FIRST_ROW As Long = 8
Private Const END_MARKER As String = "Insert new rows ABOVE this one"
Private Const DAYS_FORMULA As String = "=IF(OR(ISBLANK(task_start),ISBLANK(task_end)),"""",task_end-task_start+1)"
Private Const WHITE As Long = 16777215

' Header / task fill colours, cycled for new sections (same as the existing sections).
Private Function HeaderColors() As Variant
    HeaderColors = Array(14994616, 12379352, 12040422, 14336204, 15261367)
End Function

Private Function TaskColors() As Variant
    TaskColors = Array(15853276, 14610923, 14408946, 15523812, 15986394)
End Function

' ---------------------------------------------------------------- buttons

Public Sub AddSection()
    Dim nm As String
    If ScheduleSheet() Is Nothing Then Exit Sub
    nm = InputBox("Name of the new section:", "Add Section")
    If StrPtr(nm) = 0 Or Len(Trim$(nm)) = 0 Then Exit Sub
    InsertSection Trim$(nm)
End Sub

Public Sub AddTask()
    Dim nm As String
    If ScheduleSheet() Is Nothing Then Exit Sub
    If CurrentSection() = 0 Then
        MsgBox "Select a section or a task first. The new task goes below the selected task, " & _
               "or at the end of the selected section.", vbInformation, "Add Task"
        Exit Sub
    End If
    nm = InputBox("Name of the new task:", "Add Task", "New Task")
    If StrPtr(nm) = 0 Then Exit Sub
    If Len(Trim$(nm)) = 0 Then nm = "New Task"
    InsertTask Trim$(nm)
End Sub

Public Sub DeleteRows(Optional askFirst As Boolean = True)
    Dim target As Range, nTasks As Long, nSections As Long, msg As String
    If ScheduleSheet() Is Nothing Then Exit Sub
    Set target = RowsToDelete(nTasks, nSections)
    If target Is Nothing Then
        MsgBox "Select the task or section rows you want to delete first.", vbInformation, "Delete Rows"
        Exit Sub
    End If
    msg = "Delete " & nTasks & " task(s)"
    If nSections > 0 Then msg = msg & " and " & nSections & " section(s)"
    msg = msg & "?" & vbCrLf & vbCrLf & "This can't be undone with Ctrl+Z."
    If askFirst Then
        If MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, "Delete Rows") <> vbYes Then Exit Sub
    End If
    target.EntireRow.Delete
End Sub

' ---------------------------------------------------------------- the work behind the buttons
' (Separate from the prompts so they can be called directly, e.g. Application.Run "InsertTask", "Rough In".)

' Adds a section header plus one task, after the selected section or at the end of the list.
Public Sub InsertSection(sectionName As String)
    Dim ws As Worksheet, r As Long, h As Long, at As Long
    Dim idx As Long, tmpl As Long, startDate As Double

    Set ws = ActiveSheet

    ' After the section the selection is in, otherwise at the end of the list.
    r = ActiveCell.Row
    If r >= FIRST_ROW And r < MarkerRow(ws) Then h = SectionHeader(ws, r)
    If h > 0 Then
        at = SectionEnd(ws, h) + 1
    Else
        at = LastContentRow(ws) + 1
    End If
    If at < FIRST_ROW Then at = FIRST_ROW

    ' Copy the look of this tab's sections; a tab with none left borrows from the Template tab.
    Dim src As Worksheet
    Set src = ws
    If FindRow(ws, True) = 0 Or FindRow(ws, False) = 0 Then Set src = TemplateSheet()
    If src Is Nothing Then Set src = ws
    If FindRow(src, True) = 0 Or FindRow(src, False) = 0 Then
        MsgBox "Couldn't find an existing section and task to copy the formatting from.", vbExclamation
        Exit Sub
    End If

    idx = PickColor(ws, at)
    startDate = NextStartDate(ws, at)

    Application.ScreenUpdating = False

    InsertBlankRow ws, at
    tmpl = FindRow(src, True)
    CopyRowFormat src, tmpl, ws, at
    RecolorRow ws, at, HeaderColors()(idx)
    ws.Cells(at, "B").Value = sectionName

    InsertBlankRow ws, at + 1
    tmpl = FindRow(src, False)
    CopyRowFormat src, tmpl, ws, at + 1
    RecolorRow ws, at + 1, TaskColors()(idx)
    FillTask ws, at + 1, "New Task", startDate

    Application.ScreenUpdating = True
    ws.Cells(at + 1, "B").Select
End Sub

' Adds a task below the selected task, or at the end of the selected section.
Public Sub InsertTask(taskName As String)
    Dim ws As Worksheet, r As Long, h As Long, at As Long
    Dim tmpl As Long, startDate As Double, color As Long, i As Long

    Set ws = ActiveSheet
    h = CurrentSection()
    If h = 0 Then Exit Sub
    r = ActiveCell.Row

    If IsTask(ws, r) Then at = r + 1 Else at = SectionEnd(ws, h) + 1
    startDate = NextStartDate(ws, at)

    Application.ScreenUpdating = False

    InsertBlankRow ws, at
    ' Copy the look of the task above in the same section; an empty section borrows
    ' any task and takes the task colour that matches its header.
    Dim src As Worksheet
    Set src = ws
    If at - 1 > h Then
        tmpl = at - 1
    Else
        tmpl = FindRow(ws, False)
        If tmpl = 0 And Not TemplateSheet() Is Nothing Then
            Set src = TemplateSheet()
            tmpl = FindRow(src, False)
        End If
        For i = 0 To UBound(HeaderColors())
            If ws.Cells(h, "B").Interior.Color = HeaderColors()(i) Then color = TaskColors()(i)
        Next i
    End If
    If tmpl > 0 Then CopyRowFormat src, tmpl, ws, at
    If color > 0 Then RecolorRow ws, at, color
    FillTask ws, at, taskName, startDate

    Application.ScreenUpdating = True
    ws.Cells(at, "B").Select
End Sub

' Selected rows inside the task list; a selected section header takes its tasks with it.
' Returns Nothing when no task or section is selected.
Private Function RowsToDelete(ByRef nTasks As Long, ByRef nSections As Long) As Range
    Dim ws As Worksheet, area As Range, c As Range, target As Range, picked As Range, r As Long

    Set ws = ActiveSheet
    If TypeName(Selection) <> "Range" Then Exit Function
    If MarkerRow(ws) <= FIRST_ROW Then Exit Function
    Set picked = Intersect(Selection.EntireRow, ws.Range(ws.Rows(FIRST_ROW), ws.Rows(MarkerRow(ws) - 1)))
    If picked Is Nothing Then Exit Function

    For Each area In picked.Areas
        For Each c In area.Columns(1).Cells
            r = c.Row
            If IsHeader(ws, r) Then
                nSections = nSections + 1
                AddToRange target, ws.Range(ws.Rows(r), ws.Rows(SectionEnd(ws, r)))
            ElseIf IsTask(ws, r) Then
                AddToRange target, ws.Rows(r)
            End If
        Next c
    Next area

    If target Is Nothing Then Exit Function

    ' Areas can overlap (a section plus one of its own tasks), and Excel won't delete
    ' overlapping ranges, so rebuild the range row by row while counting.
    Set picked = Nothing
    For r = FIRST_ROW To MarkerRow(ws) - 1
        If Not Intersect(target, ws.Rows(r)) Is Nothing Then
            If IsTask(ws, r) Then nTasks = nTasks + 1
            AddToRange picked, ws.Rows(r)
        End If
    Next r
    Set RowsToDelete = picked
End Function

' ---------------------------------------------------------------- helpers

' Header row of the section the active cell is in, or 0.
Private Function CurrentSection() As Long
    Dim ws As Worksheet, r As Long
    Set ws = ActiveSheet
    r = ActiveCell.Row
    If r >= FIRST_ROW And r < MarkerRow(ws) Then CurrentSection = SectionHeader(ws, r)
End Function

Private Function ScheduleSheet() As Worksheet
    If Not IsProjectSheet(ActiveSheet) Then
        MsgBox "Open a project tab first.", vbInformation
        Exit Function
    End If
    Set ScheduleSheet = ActiveSheet
End Function

' A project tab is any worksheet laid out like the Gantt template.
Public Function IsProjectSheet(sh As Object) As Boolean
    If TypeName(sh) <> "Worksheet" Then Exit Function
    IsProjectSheet = (CStr(sh.Range("B6").Value) = "TASK")
End Function

Private Function TemplateSheet() As Worksheet
    On Error Resume Next
    Set TemplateSheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET)
End Function

Private Function MarkerRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Columns("B").Find(What:=END_MARKER, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then
        MarkerRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    Else
        MarkerRow = f.Row
    End If
End Function

Private Function IsHeader(ws As Worksheet, r As Long) As Boolean
    IsHeader = ws.Cells(r, "B").Font.Bold = True And Len(ws.Cells(r, "B").Value) > 0 _
               And IsEmpty(ws.Cells(r, "E").Value)
End Function

Private Function IsTask(ws As Worksheet, r As Long) As Boolean
    If IsHeader(ws, r) Then Exit Function
    IsTask = Len(ws.Cells(r, "B").Value) > 0 Or Not IsEmpty(ws.Cells(r, "E").Value)
End Function

' Header row of the section that row r belongs to, or 0 if r isn't in a section.
Private Function SectionHeader(ws As Worksheet, ByVal r As Long) As Long
    Do While r >= FIRST_ROW
        If IsHeader(ws, r) Then SectionHeader = r: Exit Function
        If Not IsTask(ws, r) Then Exit Function
        r = r - 1
    Loop
End Function

' Last row of the section whose header is on row h (h itself if it has no tasks).
Private Function SectionEnd(ws As Worksheet, h As Long) As Long
    Dim r As Long, last As Long
    last = MarkerRow(ws)
    r = h + 1
    Do While r < last
        If Not IsTask(ws, r) Then Exit Do
        r = r + 1
    Loop
    SectionEnd = r - 1
End Function

Private Function LastContentRow(ws As Worksheet) As Long
    Dim r As Long
    For r = MarkerRow(ws) - 1 To FIRST_ROW Step -1
        If IsHeader(ws, r) Or IsTask(ws, r) Then LastContentRow = r: Exit Function
    Next r
    LastContentRow = FIRST_ROW - 1
End Function

' First header (wantHeader = True) or first task row in the list, used as a formatting template.
Private Function FindRow(ws As Worksheet, wantHeader As Boolean) As Long
    Dim r As Long
    For r = FIRST_ROW To MarkerRow(ws) - 1
        If wantHeader And IsHeader(ws, r) Then FindRow = r: Exit Function
        If Not wantHeader And IsTask(ws, r) Then FindRow = r: Exit Function
    Next r
End Function

' Palette index for a section inserted at row `at`: the next colour in the cycle,
' skipped forward if it matches the section above or below.
Private Function PickColor(ws As Worksheet, at As Long) As Long
    Dim r As Long, n As Long, idx As Long, i As Long, prevFill As Long, nextFill As Long
    prevFill = -1: nextFill = -1
    For r = at - 1 To FIRST_ROW Step -1
        If IsHeader(ws, r) Then prevFill = ws.Cells(r, "B").Interior.Color: Exit For
    Next r
    For r = at To MarkerRow(ws) - 1
        If IsHeader(ws, r) Then nextFill = ws.Cells(r, "B").Interior.Color: Exit For
    Next r
    For r = FIRST_ROW To at - 1
        If IsHeader(ws, r) Then n = n + 1
    Next r
    idx = n Mod (UBound(HeaderColors()) + 1)
    For i = 0 To UBound(HeaderColors())
        If HeaderColors()(idx) <> prevFill And HeaderColors()(idx) <> nextFill Then Exit For
        idx = (idx + 1) Mod (UBound(HeaderColors()) + 1)
    Next i
    PickColor = idx
End Function

' Day after the task just above row `at`; failing that the latest end date above; failing that the project start.
Private Function NextStartDate(ws As Worksheet, at As Long) As Double
    Dim r As Long, best As Double
    If at - 1 >= FIRST_ROW Then
        If IsTask(ws, at - 1) And IsDate(ws.Cells(at - 1, "F").Value) Then
            NextStartDate = CDbl(ws.Cells(at - 1, "F").Value) + 1
            Exit Function
        End If
    End If
    For r = FIRST_ROW To at - 1
        If IsTask(ws, r) And IsDate(ws.Cells(r, "F").Value) Then
            If CDbl(ws.Cells(r, "F").Value) > best Then best = CDbl(ws.Cells(r, "F").Value)
        End If
    Next r
    If best > 0 Then
        NextStartDate = best + 1
    Else
        NextStartDate = CDbl(ws.Range("E2").Value)
    End If
End Function

Private Sub InsertBlankRow(ws As Worksheet, at As Long)
    ws.Rows(at).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

' Copies columns A:H (formats, the DAYS formula and the progress data bar) from one row to another.
' The chart columns keep the formatting they got from the row insert.
Private Sub CopyRowFormat(fromWs As Worksheet, fromRow As Long, ws As Worksheet, toRow As Long)
    fromWs.Range("A" & fromRow & ":H" & fromRow).Copy
    ws.Range("A" & toRow).PasteSpecial Paste:=xlPasteAllMergingConditionalFormats
    Application.CutCopyMode = False
    ws.Range("B" & toRow & ":F" & toRow).ClearContents
    ws.Range("H" & toRow).Formula = DAYS_FORMULA
    ws.Rows(toRow).RowHeight = fromWs.Rows(fromRow).RowHeight
End Sub

' Re-fills the coloured cells of columns B:H.
Private Sub RecolorRow(ws As Worksheet, ByVal r As Long, ByVal color As Long)
    Dim c As Range
    For Each c In ws.Range("B" & r & ":H" & r).Cells
        If c.Interior.Color <> WHITE Then c.Interior.Color = color
    Next c
End Sub

Private Sub FillTask(ws As Worksheet, r As Long, taskName As String, startDate As Double)
    ws.Cells(r, "B").Value = taskName
    ws.Cells(r, "D").Value = 0
    ws.Cells(r, "E").Value = startDate
    ws.Cells(r, "F").Value = startDate
End Sub

Private Sub AddToRange(ByRef target As Range, rng As Range)
    If target Is Nothing Then Set target = rng Else Set target = Union(target, rng)
End Sub
