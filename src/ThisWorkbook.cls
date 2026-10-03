Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    modCore.Initialize
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Fail
    modCore.HandleSheetChange Sh, Target
    Exit Sub
Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub
