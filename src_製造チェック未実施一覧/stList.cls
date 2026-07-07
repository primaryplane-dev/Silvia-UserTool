Option Explicit

Private Sub cmdEnd_Click()
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    Application.ThisWorkbook.Close False
End Sub

Private Sub cmdSelect_Click()
    Call subOpen
    If Not P_Regist Then Exit Sub
    If P_DATEF = 0 And P_DATET = 0 Then Exit Sub
    Call subDisplayList
End Sub

Private Sub subOpen()
    Dim obj As frmSelect
    Set obj = New frmSelect
    obj.Show
    Set obj = Nothing
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Row < Bas_List.RW_FR Then Exit Sub
    If Target.Column > 6 Then Exit Sub
    If Me.Cells(Target.Row, 1) = "" Then Exit Sub
    P_HINO = Me.Cells(Target.Row, 7)
    P_KJNO = Me.Cells(Target.Row, 8)
    P_HINM = Me.Cells(Target.Row, 3)
    P_DATE = CDate(Me.Cells(Target.Row, 1))
    If Me.Cells(Target.Row, 2) = "ビスケット分割重量" Then
        P_SYCD = ""
        P_SYNM = ""
        Call subMain2
    ElseIf InStr(Me.Cells(Target.Row, 2), "コンベアチェック表") > 0 Then
        P_SYCD = ""
        P_SYNM = ""
        P_KTCD = Me.Cells(Target.Row, 10)
            Call Bas_List4.subPrepareEditSheet
    Else
        If InStr(Me.Cells(Target.Row, 2), "備品チェック表") > 0 Then
            P_SHBU = Me.Cells(Target.Row, 9)
            P_KTCD = Me.Cells(Target.Row, 10)
            P_SYCD = ""
            P_SYNM = ""
            Call subProcessEditList
        End If
    End If
    Cancel = True
    
End Sub

