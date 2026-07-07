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

Private Sub Workbook_Open()
    
    ' 端末依存で残る行非表示状態を初期化
    stList.Cells.EntireRow.Hidden = False

    'ボタンのサイズ調整
    stList.OLEObjects("cmdSelect").Height = stList.Cells(1, 10).Height: stList.OLEObjects("cmdSelect").Width = 80: stList.OLEObjects("cmdSelect").Left = stList.Cells(1, 10).Left: stList.OLEObjects("cmdSelect").Top = stList.Cells(1, 10).Top
    stList.OLEObjects("cmdSYCD").Height = stList.Cells(2, 10).Height: stList.OLEObjects("cmdSYCD").Width = 80: stList.OLEObjects("cmdSYCD").Left = stList.Cells(2, 10).Left: stList.OLEObjects("cmdSYCD").Top = stList.Cells(2, 10).Top
    stList.OLEObjects("cmdNew").Height = stList.Cells(1, 12).Height: stList.OLEObjects("cmdNew").Width = 80: stList.OLEObjects("cmdNew").Left = stList.Cells(1, 12).Left: stList.OLEObjects("cmdNew").Top = stList.Cells(1, 12).Top
    stList.OLEObjects("cmdSEnd").Height = stList.Cells(2, 12).Height: stList.OLEObjects("cmdSEnd").Width = 80: stList.OLEObjects("cmdSEnd").Left = stList.Cells(2, 12).Left: stList.OLEObjects("cmdSEnd").Top = stList.Cells(2, 12).Top
    stList.OLEObjects("cmdSave").Height = stList.Cells(1, 14).Height: stList.OLEObjects("cmdSave").Width = 80: stList.OLEObjects("cmdSave").Left = stList.Cells(1, 14).Left: stList.OLEObjects("cmdSave").Top = stList.Cells(1, 14).Top
    stList.OLEObjects("cmdEnd").Height = stList.Cells(2, 14).Height: stList.OLEObjects("cmdEnd").Width = 80: stList.OLEObjects("cmdEnd").Left = stList.Cells(2, 14).Left: stList.OLEObjects("cmdEnd").Top = stList.Cells(2, 14).Top
    stList.OLEObjects("cmdUp").Height = stList.Cells(1, 16).Height: stList.OLEObjects("cmdUp").Width = stList.Cells(1, 16).Width: stList.OLEObjects("cmdUp").Left = stList.Cells(1, 16).Left: stList.OLEObjects("cmdUp").Top = stList.Cells(1, 16).Top
    stList.OLEObjects("cmdDown").Height = stList.Cells(2, 16).Height: stList.OLEObjects("cmdDown").Width = stList.Cells(2, 16).Width: stList.OLEObjects("cmdDown").Left = stList.Cells(2, 16).Left: stList.OLEObjects("cmdDown").Top = stList.Cells(2, 16).Top
    
    '担当者選択
    Call subOpen
    stList.Cells(6, 8) = P_SYNM
    stList.Cells(4, CL_SY) = P_SYCD

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    stList.Cells(1, 1).Copy
    Application.CutCopyMode = False
    ThisWorkbook.Saved = True
    If Application.Workbooks.Count = 1 Then Application.Quit
End Sub

Private Sub subOpen()
    Dim obj As New frmSYCD
    obj.Show
    Set obj = Nothing
End Sub


