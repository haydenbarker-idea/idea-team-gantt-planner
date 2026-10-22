Attribute VB_Name = "SnapshotTools"
Option Explicit

' Buttons for the project tabs: Export Snapshot and Email Snapshot to Client.
' A snapshot is a customer copy of one project tab: its own workbook with values only
' (frozen as of today), no buttons, links or invoice markers, the chart cropped to the
' project's dates, and a page setup that prints one page wide. It is saved as a PDF and
' as an .xlsx; the email version attaches both to an Outlook draft for the client.

Private Const PROJECTS_SHEET As String = "Projects"
Private Const FIRST_ROW As Long = 8
Private Const END_MARKER As String = "Insert new rows ABOVE this one"
Private Const CHART_FIRST_COL As Long = 9     ' column I
Private Const CHART_LAST_COL As Long = 134    ' column ED
Private Const CLIENT_EMAIL_COL As String = "K"

Public SnapshotError As String                ' why the last snapshot failed, for messages
Private stage As String

' ---------------------------------------------------------------- buttons

Public Sub ExportSnapshot()
    Dim src As Worksheet, folder As Variant, files As Variant
    Set src = ProjectTab()
    If src Is Nothing Then Exit Sub

    folder = PickFolder()
    If Len(folder) = 0 Then Exit Sub

    files = SaveSnapshot(src, CStr(folder))
    If IsEmpty(files) Then
        MsgBox "The snapshot couldn't be created (" & SnapshotError & ").", vbExclamation, "Export Snapshot"
        Exit Sub
    End If
    MsgBox "Saved:" & vbCrLf & files(0) & vbCrLf & files(1), vbInformation, "Export Snapshot"
End Sub

Public Sub EmailSnapshot()
    Dim src As Worksheet
    Set src = ProjectTab()
    If src Is Nothing Then Exit Sub
    DraftSnapshotEmail src, True
End Sub

' ---------------------------------------------------------------- the work behind the buttons

' Saves "<project> - Schedule <date>.pdf" and ".xlsx" for the project tab into folder.
' Returns Array(pdfPath, xlsxPath), or Empty if something went wrong (see SnapshotError).
Public Function SaveSnapshot(src As Worksheet, ByVal folder As String) As Variant
    Dim wb As Workbook, ws As Worksheet, base As String, pdfPath As String, xlsxPath As String

    SnapshotError = ""
    On Error GoTo Failed
    stage = "naming the files"
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    base = folder & SafeFileName(CStr(src.Range("B1").Value)) & " - Schedule " & Format$(Date, "yyyy-mm-dd")
    pdfPath = base & ".pdf"
    xlsxPath = base & ".xlsx"

    Application.ScreenUpdating = False
    stage = "copying the tab"
    src.Copy                                   ' new workbook holding just this tab
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Unprotect                               ' the live tabs are protected; the copy is reworked, then frozen
    ws.Name = "Schedule"
    MakeCustomerView ws

    stage = "saving the PDF"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    stage = "saving the Excel copy"
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=xlsxPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    src.Activate
    Application.ScreenUpdating = True

    SaveSnapshot = Array(pdfPath, xlsxPath)
    Exit Function

Failed:
    SnapshotError = stage & ": " & Err.Description
    SetPrintCommunication True
    Application.DisplayAlerts = True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Function

Public Function SnapshotErrorText() As String
    SnapshotErrorText = SnapshotError
End Function

' Same as SaveSnapshot, by tab name (for Application.Run and automation).
Public Function SaveSnapshotFor(tabName As String, folder As String) As Variant
    SaveSnapshotFor = SaveSnapshot(ThisWorkbook.Worksheets(tabName), folder)
End Function

' Same as DraftSnapshotEmail, by tab name (for Application.Run and automation).
Public Sub DraftSnapshotEmailFor(tabName As String, showDraft As Boolean)
    DraftSnapshotEmail ThisWorkbook.Worksheets(tabName), showDraft
End Sub

' Creates an Outlook draft to the project's client email with both snapshot files attached.
' showDraft = True opens it for the user to review and send; False just saves it to Drafts.
Public Sub DraftSnapshotEmail(src As Worksheet, Optional ByVal showDraft As Boolean = True)
    Dim folder As String, files As Variant, ol As Object, mail As Object
    Dim projectName As String, intro As String

    folder = Environ$("TEMP") & "\Project Snapshots\"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    files = SaveSnapshot(src, folder)
    If IsEmpty(files) Then
        If showDraft Then MsgBox "The snapshot couldn't be created (" & SnapshotError & ").", vbExclamation, "Email Snapshot"
        Exit Sub
    End If

    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        MsgBox "Couldn't start Outlook. The files were saved here instead:" & vbCrLf & folder, vbExclamation
        Exit Sub
    End If

    projectName = CStr(src.Range("B1").Value)
    Set mail = ol.CreateItem(0)
    mail.To = ClientEmail(src.Name)
    mail.Subject = projectName & " - schedule update " & Format$(Date, "mmm d, yyyy")
    mail.Attachments.Add CStr(files(0))
    mail.Attachments.Add CStr(files(1))
    intro = "<p>Hi,</p><p>Please find attached the latest schedule for <b>" & HtmlText(projectName) & _
            "</b>, as of " & Format$(Date, "mmmm d, yyyy") & ", as a PDF and an Excel copy.</p>" & _
            "<p>Let me know if you have any questions.</p>"
    If showDraft Then
        mail.Display                             ' adds the user's default signature
        mail.HTMLBody = intro & mail.HTMLBody
    Else
        mail.HTMLBody = intro
        mail.Save
    End If
End Sub

' ---------------------------------------------------------------- helpers

' Turns the copied tab into the customer view: values only, internal bits removed,
' chart cropped to the project, print-ready.
Private Sub MakeCustomerView(ws As Worksheet)
    Dim i As Long, marker As Long, lastRow As Long, r As Long
    Dim lastEnd As Double, startDay As Double, nDays As Long, lastCol As Long, width As Double
    Dim f As Range

    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Type <> msoComment Then ws.Shapes(i).Delete
    Next i
    For i = ws.Comments.Count To 1 Step -1
        ws.Comments(i).Delete
    Next i
    ws.Range("I1").MergeArea.ClearContents       ' "All projects" link
    ws.Range("I3:ED3").ClearContents             ' invoice markers (internal)
    ws.Range("E4").Value = 1                     ' chart starts at the project's first week
    ws.Calculate
    ws.UsedRange.Value = ws.UsedRange.Value      ' freeze as of today, no formulas
    ws.Range("D3").Value = "As of:"
    ws.Range("D4:E4").Clear                      ' Display Week is an internal control
    ws.Range(ws.Cells(5, CHART_FIRST_COL), ws.Cells(5, CHART_LAST_COL)).Interior.Color = _
        ws.Cells(4, CHART_FIRST_COL).Interior.Color   ' drop hand-coloured (internal) day markers

    Set f = ws.Columns("B").Find(What:=END_MARKER, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then marker = ws.UsedRange.Rows.Count + 1 Else marker = f.Row
    lastRow = FIRST_ROW
    For r = FIRST_ROW To marker - 1
        If Len(ws.Cells(r, "B").Value) > 0 Then lastRow = r
        If IsDate(ws.Cells(r, "F").Value) Then
            If CDbl(ws.Cells(r, "F").Value) > lastEnd Then lastEnd = CDbl(ws.Cells(r, "F").Value)
        End If
    Next r

    ' Whole weeks up to the last end date, within the chart's 18 weeks.
    startDay = CDbl(ws.Cells(5, CHART_FIRST_COL).Value)
    nDays = 7
    If lastEnd >= startDay Then nDays = Application.WorksheetFunction.Ceiling(lastEnd - startDay + 1, 7)
    lastCol = CHART_FIRST_COL + nDays - 1
    If lastCol > CHART_LAST_COL Then lastCol = CHART_LAST_COL

    If lastCol < CHART_LAST_COL Then ws.Range(ws.Columns(lastCol + 1), ws.Columns(CHART_LAST_COL)).Delete
    If ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 > lastRow Then
        ws.Range(ws.Rows(lastRow + 1), ws.Rows(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)).Delete
    End If

    ' Shorter dates so the task columns take less of the page.
    ws.Range("E" & FIRST_ROW & ":F" & lastRow).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns("E:F").ColumnWidth = 17

    For i = 1 To lastCol
        width = width + ws.Columns(i).Width
    Next i

    stage = "setting up the page"
    SetPrintCommunication False                  ' faster page setup, where the printer allows it
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
        .PrintTitleRows = "$4:$6"
        .Orientation = xlLandscape
        .PaperSize = IIf(width > 1100, xlPaperTabloid, xlPaperLetter)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.6)
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "Idea Networks"
        .CenterFooter = "Page &P of &N"
        .RightFooter = "Schedule as of " & Format$(Date, "mmm d, yyyy")
        .PrintGridlines = False
    End With
    SetPrintCommunication True

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

' Some printer set-ups refuse PrintCommunication; page setup still works without it, just slower.
Private Sub SetPrintCommunication(ByVal value As Boolean)
    On Error Resume Next
    Application.PrintCommunication = value
End Sub

Private Function ProjectTab() As Worksheet
    If TypeName(ActiveSheet) = "Worksheet" Then
        If CStr(ActiveSheet.Range("B6").Value) = "TASK" And ActiveSheet.Name <> "Template" Then
            Set ProjectTab = ActiveSheet
            Exit Function
        End If
    End If
    MsgBox "Open the project's tab first.", vbInformation
End Function

Private Function PickFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Save the snapshot PDF and Excel file to"
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

' Client email from the Projects list (column K), matched on the Tab column (J).
Private Function ClientEmail(tabName As String) As String
    Dim list As Worksheet, r As Long
    On Error Resume Next
    Set list = ThisWorkbook.Worksheets(PROJECTS_SHEET)
    On Error GoTo 0
    If list Is Nothing Then Exit Function
    For r = 6 To list.Cells(list.Rows.Count, "J").End(xlUp).Row
        If StrComp(CStr(list.Cells(r, "J").Value), tabName, vbTextCompare) = 0 Then
            ClientEmail = Trim$(CStr(list.Cells(r, CLIENT_EMAIL_COL).Value))
            Exit Function
        End If
    Next r
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, " ")
    Next ch
    SafeFileName = Trim$(s)
    If Len(SafeFileName) = 0 Then SafeFileName = "Project"
End Function

Private Function HtmlText(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    HtmlText = Replace(s, ">", "&gt;")
End Function
