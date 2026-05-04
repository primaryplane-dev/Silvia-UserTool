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
    
    '「アイテム選択」（ビスケット、クッキー、ドーナツ）
    'SBFP01にはSHBU が無いので選択画面はスキップ。
    'SHBUが必要可能性もあるのでコメントアウト対応。

    'ActiveXコントロールボタン廃止のため、サイズ調整処理は不要
    
    
    'Call subOpenSelectSHBU
    'If P_CalendarSelected And Not P_SHBU = "" Then
        Call subMain
    'End If
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    stList.Cells(1, 1).Copy
    Application.CutCopyMode = False
    If Workbooks.Count = 1 Then Application.Quit
    ThisWorkbook.Saved = True
End Sub

Private Sub subOpenSelectSHBU()
    Dim obj     As New frmSelectSHBU
    obj.Show
    Set obj = Nothing
End Sub


