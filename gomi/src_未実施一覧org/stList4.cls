VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stList4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdSYCD_Click()
    Call subOpenSYCD
    Me.Cells(9, 2).Select
End Sub

Private Sub subOpenSYCD()
    Dim obj As New frmSYCD
    obj.Show
    Set obj = Nothing
End Sub

Private Sub cmdSave_Click()
    If Me.Cells(1, 14) = "" Then Exit Sub
    If Me.Cells(5, 3) = "" Or Me.Cells(5, 5) = "" Then Exit Sub
    Call subUpdate4
    Call subMain4
End Sub

Private Sub cmdBack_Click()
    If Me.Cells(1, 14) = "1" Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Call subMain
End Sub

Private Sub cmdUp_Click()
    If ActiveWindow.ScrollRow < 12 Then Exit Sub
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 3
End Sub

Private Sub cmdDown_Click()
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 3
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Target.Row = 7 Then Exit Sub
    If Target.Column < 11 Or Target.Column > 13 Then Exit Sub
    If Me.Cells(5, 3) = "" Or Me.Cells(5, 5) = "" Then Exit Sub
    If P_SYCD = "" Then MsgBox ("担当者を選択してください"): Exit Sub
    Dim strSYNM     As String
    
    '最終確認者、品管、ライン長
    If Not Me.Cells(1, 4) = "製造終了" Then Exit Sub
    If Target.Column = 11 Then                              '最終確認者
        If Me.Cells(7, 12) = "" Or Me.Cells(7, 13) = "" Then Exit Sub
        If Me.Cells(Target.Row, Target.Column) = "" Then
            strSYNM = fncChkK(3)
            If strSYNM = "" Then MsgBox ("権限がありません"): Exit Sub
            Me.Cells(Target.Row, Target.Column) = strSYNM
            Me.Cells(Target.Row - 1, Target.Column) = P_SYCD
        End If
    ElseIf Target.Column = 12 Or Target.Column = 13 Then    '品管,ライン長
        If Me.Cells(Target.Row, Target.Column) = "" Then
            strSYNM = fncChkK(2)
            If strSYNM = "" Then MsgBox ("権限がありません"): Exit Sub
            Me.Cells(Target.Row, Target.Column) = strSYNM
            Me.Cells(Target.Row - 1, Target.Column) = P_SYCD
        End If
    End If
    Me.Cells(1, 14) = "1"
    Me.Cells(9, 2).Select
End Sub

