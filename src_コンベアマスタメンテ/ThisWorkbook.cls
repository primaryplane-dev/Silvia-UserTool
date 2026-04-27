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

    'ボタンのサイズ調整
'    With stList.OLEObjects("cmdSelect")　'「アイテム選択」ボタンは非表示処置
'        .Height = 27: .Width = 90: .Top = 9: .Left = 219
'    End With
    With stList.OLEObjects("cmdUpdate")
        .Height = 27: .Width = 90: .Top = 9: .Left = 318
    End With
    With stList.OLEObjects("cmdEnd")
        .Height = 27: .Width = 90: .Top = 9: .Left = 417
    End With
    
    
    'Call subOpenSelectSHBU
    'If P_Regist2 And Not P_SHBU = "" Then
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


