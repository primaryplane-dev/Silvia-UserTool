Option Explicit

Private Sub cmdSYCD_Click()
    Call subOpenSYCD

    Me.Cells(9, 1).Select
End Sub

Private Sub subOpenSYCD()
    Dim obj As frmSYCD
    Set obj = New frmSYCD
    obj.Show
    Set obj = Nothing
End Sub

Private Sub cmdSave_Click()
    If Me.Cells(6, 15) = "" Then Exit Sub
    If Me.Cells(10, 6) = "" Then Exit Sub
    Call subUpdate2
    Call subMain2
End Sub

Private Sub cmdBack_Click()
    If Not Me.Cells(6, 15) = "" Then
        If MsgBox("ï€ë∂Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÇ™ÅAÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Call subShowList
End Sub

Private Sub cmdUp_Click()
    If ActiveWindow.ScrollRow < 10 Then Exit Sub
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 3
End Sub

Private Sub cmdDown_Click()
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 3
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Target.Row = 5 Then Exit Sub
    If Target.Column < 12 And Target.Column > 14 Then Exit Sub
    
    If P_SYCD = "" Then MsgBox ("íSìñé“ÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢"): Exit Sub
    Dim KBN     As Integer
    Dim strSYNM As String
    Dim lRow    As Long
        
    If Target.Row = 5 Then
        If Me.Cells(1, 8) = "ê∂éYíÜ" Then Exit Sub
        If Me.Cells(Target.Row, Target.Column) = "" Then
            If Target.Column = 12 Then
                If Me.Cells(Target.Row, 13) = "" Or Me.Cells(Target.Row, 14) = "" Then Exit Sub
                KBN = 3
            End If
            If Target.Column = 13 Or Target.Column = 14 Then KBN = 2
            strSYNM = fncChkK(KBN)
            If strSYNM = "" Then MsgBox ("å†å¿Ç™Ç†ÇËÇ‹ÇπÇÒ"): Exit Sub
            Me.Cells(Target.Row, Target.Column) = strSYNM
            Me.Cells(Target.Row + 1, Target.Column) = P_SYCD
            Me.Cells(6, 15) = "1"
        End If
    End If
End Sub

