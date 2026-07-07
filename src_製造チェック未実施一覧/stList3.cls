VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stList3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdSYCD_Click()
    Call subOpenSYCD
    Me.Cells(8, 1).Select
End Sub

Private Sub subOpenSYCD()
    Dim obj As frmSYCD
    Set obj = New frmSYCD
    obj.Show
    Set obj = Nothing
End Sub

Private Sub cmdSave_Click()
    If Me.Cells(1, 17) = "" Then Exit Sub
    If Me.Cells(4, 5) = "" Or Me.Cells(4, 8) = "" Then Exit Sub
    Call subUpdate3
    Call subProcessEditList
End Sub

Private Sub cmdBack_Click()
    If Me.Cells(1, 17) = "1" Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Call subDisplayList
End Sub

Private Sub cmdUp_Click()
    If ActiveWindow.ScrollRow < 11 Then Exit Sub
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 3
End Sub

Private Sub cmdDown_Click()
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 3
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Target.Row = 6 Then Exit Sub
    If Target.Column < 14 Or Target.Column > 16 Then Exit Sub
    If Me.Cells(4, 5) = "" Or Me.Cells(4, 8) = "" Then Exit Sub
    If P_SYCD = "" Then MsgBox ("担当者を選択してください"): Exit Sub
    Dim strSYNM     As String
    
    '最終確認者、品管、ライン長
    If Not Me.Cells(1, 7) = "製造終了" Then Exit Sub
    If Target.Column = 14 Then                              '最終確認者
        If Me.Cells(6, 15) = "" Or Me.Cells(6, 16) = "" Then Exit Sub
        If Me.Cells(Target.Row, Target.Column) = "" Then
            strSYNM = fncChkK(3)
            If strSYNM = "" Then MsgBox ("権限がありません"): Exit Sub
            Me.Cells(Target.Row, Target.Column) = strSYNM
            Me.Cells(Target.Row - 1, Target.Column) = P_SYCD
        End If
    ElseIf Target.Column = 15 Or Target.Column = 16 Then    '品管,ライン長
        If Me.Cells(Target.Row, Target.Column) = "" Then
            strSYNM = fncChkK(2)
            If strSYNM = "" Then MsgBox ("権限がありません"): Exit Sub
            Me.Cells(Target.Row, Target.Column) = strSYNM
            Me.Cells(Target.Row - 1, Target.Column) = P_SYCD
        End If
    End If
    Me.Cells(1, 17) = "1"
    Me.Cells(8, 1).Select
End Sub

