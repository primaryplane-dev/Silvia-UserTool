VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    stList.Cells(1, 1).Copy
    Application.CutCopyMode = False
    ThisWorkbook.Saved = True
    If Application.Workbooks.Count = 1 Then Application.Quit
End Sub

