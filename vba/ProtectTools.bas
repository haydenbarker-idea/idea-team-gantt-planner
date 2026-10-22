Attribute VB_Name = "ProtectTools"
Option Explicit

' Sheet protection against accidental edits (no password), the same as the online version.
' Which cells stay editable is stored in the file (each cell's Locked flag):
'   project tabs + Template: B1 name, B3 manager, E2 start, E4 Display Week, row 3 chart labels
'     and whole task rows (7 down to just above the marker row), so rows can still be deleted.
'   Projects: Client (C), Tab (J) and Client Email (K).   Instructions: nothing.
' UserInterfaceOnly lets the macros change locked cells, but Excel forgets it when the file
' closes, so it's reapplied when the workbook opens and before each button's work.

Private Const PROJECTS_SHEET As String = "Projects"
Private Const INSTRUCTIONS_SHEET As String = "Instructions"

Public Sub ProtectAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsProjectSheet(ws) Then
            ProtectSchedule ws
        ElseIf ws.Name = PROJECTS_SHEET Then
            ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
                AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
        ElseIf ws.Name = INSTRUCTIONS_SHEET Then
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        End If
    Next ws
End Sub

' Project tabs: rows can be inserted and deleted and anything formatted; columns can't be added or removed.
Public Sub ProtectSchedule(ws As Worksheet)
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
        AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowDeletingRows:=True
End Sub
