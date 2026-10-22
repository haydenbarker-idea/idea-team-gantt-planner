Attribute VB_Name = "ProjectTools"
Option Explicit

' Buttons for the Projects tab: New Project and Archive / Restore.
' Each project lives on its own tab, copied from the hidden Template tab. The Projects
' tab lists them from row LIST_FIRST_ROW down, one row per project; its columns are
'   B Project (link)  C Client  D Project manager  E Start  F Finish
'   G Progress  H Status  I Archived ("Yes" or blank)  J Tab name  K Client email
' Everything but Client, Archived, Tab name and Client email is a formula reading the project's tab;
' its ranges start on the blank row 7 so rows added at the top of the task list are included.

Private Const PROJECTS_SHEET As String = "Projects"
Private Const TEMPLATE_SHEET As String = "Template"
Private Const LIST_FIRST_ROW As Long = 6
Private Const END_MARKER As String = "Insert new rows ABOVE this one"

' ---------------------------------------------------------------- buttons

Public Sub NewProject()
    Dim nm As String, client As String, email As String, pm As String, startText As String

    nm = InputBox("Project name (example: Atlantic Packaging - Line 3):", "New Project")
    If StrPtr(nm) = 0 Or Len(Trim$(nm)) = 0 Then Exit Sub
    client = InputBox("Client (example: Atlantic Packaging), optional:", "New Project")
    If StrPtr(client) = 0 Then Exit Sub
    email = InputBox("Client email, for Email Snapshot to Client (example: jane.doe@example.com), optional:", "New Project")
    If StrPtr(email) = 0 Then Exit Sub
    pm = InputBox("Project manager (example: Vinny Singh), optional:", "New Project")
    If StrPtr(pm) = 0 Then Exit Sub
    startText = InputBox("Start date (example: 2026-11-02 or 11/2/2026). Today is filled in:", "New Project", Format$(Date, "yyyy-mm-dd"))
    If StrPtr(startText) = 0 Then Exit Sub
    If Not IsDate(startText) Then
        MsgBox "'" & startText & "' isn't a date. Try again with a date like " & Format$(Date, "yyyy-mm-dd") & ".", vbExclamation
        Exit Sub
    End If

    CreateProject Trim$(nm), Trim$(client), Trim$(pm), CDate(startText), Trim$(email)
End Sub

' Archives the selected project (on the Projects tab) or the project tab you're on:
' marks it Archived and hides its tab. On an archived project it restores it instead.
Public Sub ArchiveProject()
    Dim list As Worksheet, r As Long
    Set list = ThisWorkbook.Worksheets(PROJECTS_SHEET)

    If ActiveSheet.Name = PROJECTS_SHEET Then
        r = ActiveCell.Row
        If r < LIST_FIRST_ROW Or Len(list.Cells(r, "J").Value) = 0 Then r = 0
    Else
        r = ListRow(ActiveSheet.Name)
    End If
    If r = 0 Then
        MsgBox "Select a project in the list first.", vbInformation, "Archive / Restore"
        Exit Sub
    End If

    SetArchived r, list.Cells(r, "I").Value <> "Yes"
End Sub

' ---------------------------------------------------------------- the work behind the buttons

' Copies the Template tab into a new project tab and adds it to the Projects list.
Public Sub CreateProject(projectName As String, client As String, pm As String, startDate As Date, _
                         Optional clientEmail As String = "")
    Dim tmpl As Worksheet, ws As Worksheet, wasVisible As XlSheetVisibility, r As Long

    Set tmpl = ThisWorkbook.Worksheets(TEMPLATE_SHEET)
    Application.ScreenUpdating = False

    wasVisible = tmpl.Visible
    tmpl.Visible = xlSheetVisible
    tmpl.Copy Before:=tmpl
    Set ws = ActiveSheet  ' the copy
    tmpl.Visible = wasVisible
    ws.Name = UniqueSheetName(projectName)

    ws.Range("B1").Value = projectName
    ws.Range("B3").Value = pm
    ws.Range("E2").Value = startDate
    ws.Range("E4").Value = 1
    ' The template's tasks all start on the project start date.
    For r = 8 To ws.Columns("B").Find(What:=END_MARKER, LookIn:=xlValues, LookAt:=xlWhole).Row - 1
        If Not IsEmpty(ws.Cells(r, "E").Value) Then
            ws.Cells(r, "E").Value = startDate
            ws.Cells(r, "F").Value = startDate
        End If
    Next r

    AddToList ws.Name, client, clientEmail

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("B8").Select
End Sub

Public Sub SetArchived(listRow As Long, archive As Boolean)
    Dim list As Worksheet, ws As Worksheet
    Set list = ThisWorkbook.Worksheets(PROJECTS_SHEET)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CStr(list.Cells(listRow, "J").Value))
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Couldn't find the tab '" & list.Cells(listRow, "J").Value & "'. If it was renamed, " & _
               "put the new name in the Tab column.", vbExclamation
        Exit Sub
    End If

    If archive Then
        list.Activate
        list.Cells(listRow, "I").Value = "Yes"
        ws.Visible = xlSheetHidden
    Else
        list.Cells(listRow, "I").ClearContents
        ws.Visible = xlSheetVisible
        ws.Activate
    End If
End Sub

' ---------------------------------------------------------------- helpers

Private Sub AddToList(tabName As String, client As String, clientEmail As String)
    Dim list As Worksheet, r As Long, q As String
    Set list = ThisWorkbook.Worksheets(PROJECTS_SHEET)

    r = list.Cells(list.Rows.Count, "J").End(xlUp).Row + 1
    If r < LIST_FIRST_ROW Then r = LIST_FIRST_ROW
    If r > LIST_FIRST_ROW Then
        list.Rows(r - 1).Copy
        list.Rows(r).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    q = "'" & Replace(tabName, "'", "''") & "'!"
    list.Cells(r, "J").Value = tabName
    list.Cells(r, "B").Formula = "=HYPERLINK(""#'""&SUBSTITUTE(J" & r & ",""'"",""''"")&""'!A1""," & q & "$B$1)"
    list.Cells(r, "C").Value = client
    list.Cells(r, "K").Value = clientEmail
    list.Cells(r, "D").Formula = "=" & q & "$B$3&"""""
    list.Cells(r, "E").Formula = "=IF(COUNT(" & q & "$E$7:$E$1000)=0,""""," & "MIN(" & q & "$E$7:$E$1000))"
    list.Cells(r, "F").Formula = "=IF(COUNT(" & q & "$F$7:$F$1000)=0,""""," & "MAX(" & q & "$F$7:$F$1000))"
    list.Cells(r, "G").Formula = "=IFERROR(SUMPRODUCT(" & q & "$D$7:$D$1000," & q & "$H$7:$H$1000)/SUM(" & q & "$H$7:$H$1000),"""")"
    list.Cells(r, "H").Formula = "=IF(I" & r & "=""Yes"",""Archived"",IF(E" & r & "="""",""Not scheduled""," & _
        "IF(N(G" & r & ")>=1,""Complete"",IF(TODAY()<E" & r & ",""Not started""," & _
        "IF(TODAY()>F" & r & ",""Overdue"",""In progress"")))))"
End Sub

Private Function ListRow(tabName As String) As Long
    Dim list As Worksheet, r As Long
    Set list = ThisWorkbook.Worksheets(PROJECTS_SHEET)
    For r = LIST_FIRST_ROW To list.Cells(list.Rows.Count, "J").End(xlUp).Row
        If StrComp(list.Cells(r, "J").Value, tabName, vbTextCompare) = 0 Then ListRow = r: Exit Function
    Next r
End Function

' Tab names: max 31 characters, none of : \ / ? * [ ], and unique in the workbook.
Private Function UniqueSheetName(projectName As String) As String
    Dim base As String, nm As String, ch As Variant, n As Long
    base = projectName
    For Each ch In Array(":", "\", "/", "?", "*", "[", "]")
        base = Replace(base, ch, " ")
    Next ch
    Do While InStr(base, "  ") > 0
        base = Replace(base, "  ", " ")
    Loop
    base = Trim$(base)
    If Len(base) = 0 Then base = "Project"
    base = Trim$(Left$(base, 31))
    nm = base
    n = 1
    Do While SheetExists(nm)
        n = n + 1
        nm = Left$(base, 31 - Len(" (" & n & ")")) & " (" & n & ")"
    Loop
    UniqueSheetName = nm
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim sh As Object
    For Each sh In ThisWorkbook.Sheets
        If StrComp(sh.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next sh
End Function
