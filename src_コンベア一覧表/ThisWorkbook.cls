Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    stList.Cells(1, 1).Copy
    Application.CutCopyMode = False
    ThisWorkbook.Saved = True
    If Application.Workbooks.Count = 1 Then Application.Quit
End Sub
