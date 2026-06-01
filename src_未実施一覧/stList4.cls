Option Explicit

Private mPendingCol As Long

Private Sub cmdSYCD_Click()
    Call subOpenSYCD
    If P_SYCD <> "" Then
        If mPendingCol >= 11 And mPendingCol <= 13 Then
            Call subApplyApproverByCol(mPendingCol)
            mPendingCol = 0
        End If
    End If
    Me.Cells(9, 2).Select
End Sub

Private Sub subOpenSYCD()
    Dim obj As frmSYCD
    Set obj = New frmSYCD
    obj.Show
    Set obj = Nothing
End Sub

Private Sub cmdSave_Click()
    If Me.Cells(1, 14) = "" Then Exit Sub
    If Me.Cells(5, 3) = "" Or Me.Cells(5, 5) = "" Then Exit Sub
    Call subUpdate4
    Call subPrepareEditSheet
    Me.Cells(1, 14) = ""
End Sub

Private Sub cmdBack_Click()
    If Val(Me.Cells(1, 14).Value) = 1 Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Call subDisplayList
End Sub

Private Sub cmdUp_Click()
    If ActiveWindow.ScrollRow < 12 Then Exit Sub
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 3
End Sub

Private Sub cmdDown_Click()
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 3
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim lRow        As Long
    Dim lCol        As Long
    Dim lInputCol   As Long

    lRow = Target.Row
    lCol = Target.Column

    'L/Mが結合セル等でL列判定になる場合の補正
    If lCol = 12 Then
        If Me.Cells(7, 12) <> "" And Me.Cells(7, 13) = "" Then
            lCol = 13
        End If
    End If

    If lRow <> 6 And lRow <> 7 Then Exit Sub
    If lCol < 11 Or lCol > 13 Then Exit Sub
    If Me.Cells(5, 3) = "" Or Me.Cells(5, 5) = "" Then Exit Sub

    If lCol = 12 Or lCol = 13 Then
        lInputCol = lCol
        '結合セル等でL/Mの判定が不安定な場合でも、未入力側へ確実に設定する
        If lInputCol = 12 And Me.Cells(7, 13) = "" Then lInputCol = 13
        If lInputCol = 13 And Me.Cells(7, 13) <> "" And Me.Cells(7, 12) = "" Then lInputCol = 12
    Else
        lInputCol = lCol
    End If

    If P_SYCD = "" Then
        mPendingCol = lInputCol
        MsgBox ("担当者を選択してください")
        Exit Sub
    End If

    Call subApplyApproverByCol(lInputCol)
End Sub

Private Sub subApplyApproverByCol(ByVal lCol As Long)
    Dim strSYNM     As String
    
    '最終確認者、品管、ライン長
    If Not Me.Cells(1, 4) = "製造終了" Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo Exit_

    If lCol = 11 Then                              '最終確認者
        strSYNM = fncChkK(3)
        If strSYNM = "" Then strSYNM = fncChkK(2)
        If strSYNM = "" Then strSYNM = P_SYNM
        If strSYNM = "" Then MsgBox ("担当者を選択してください"): Exit Sub
        If Me.Cells(7, 12) = "" Or Me.Cells(7, 13) = "" Then
            MsgBox ("先にL7(品管)とM7(ライン長)を入力してください")
            Exit Sub
        End If
        Me.Cells(7, lCol) = strSYNM
        Me.Cells(6, lCol) = P_SYCD
    ElseIf lCol = 12 Or lCol = 13 Then    '品管,ライン長
        strSYNM = fncChkK(2)
        If strSYNM = "" Then strSYNM = P_SYNM
        If strSYNM = "" Then MsgBox ("担当者を選択してください"): Exit Sub
        Me.Cells(7, lCol) = strSYNM
        Me.Cells(6, lCol) = P_SYCD
    End If
    Me.Cells(1, 14) = "1"
    Me.Cells(9, 2).Select

Exit_:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Call Worksheet_SelectionChange(Target)
    If (Target.Row = 6 Or Target.Row = 7) And Target.Column >= 11 And Target.Column <= 13 Then
        Cancel = True
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngCode As Range
    Dim rngName As Range
    Dim c As Range
    Dim sCode As String
    Dim sName As String
    Dim sSYCD As String

    Set rngCode = Intersect(Target, Me.Range(Me.Cells(6, 11), Me.Cells(6, 13)))
    Set rngName = Intersect(Target, Me.Range(Me.Cells(7, 11), Me.Cells(7, 13)))
    If rngCode Is Nothing And rngName Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo Exit_

    If Not rngCode Is Nothing Then
        For Each c In rngCode.Cells
            sCode = Trim$(CStr(c.Value))
            If Not (sCode = "" Or Val(sCode) = 0) Then
                sName = fncGetSYNMByCode(sCode)
                If sName <> "" Then Me.Cells(7, c.Column).Value = sName
            End If
        Next c
    End If

    If Not rngName Is Nothing Then
        For Each c In rngName.Cells
            sName = Trim$(CStr(c.Value))
            If sName <> "" Then
                sSYCD = fncGetSYCDByName(sName)
                If sSYCD <> "" Then Me.Cells(6, c.Column).Value = sSYCD
            End If
        Next c
    End If

    Me.Cells(1, 14).Value = "1"

Exit_:
    Application.EnableEvents = True
End Sub

Private Function fncGetSYNMByCode(ByVal iSYCD As String) As String
    Dim CN          As Object
    Dim RS          As Object
    Dim strSQL      As String

    fncGetSYNMByCode = ""
    If Val(iSYCD) = 0 Then Exit Function

    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")

    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString

    strSQL = ""
    strSQL = strSQL & "SELECT VASYKJ "
    strSQL = strSQL & "  FROM LIBBMF.BVAP01 "
    strSQL = strSQL & " WHERE VAKYUK='' "
    strSQL = strSQL & "   AND VASYCD=" & Val(iSYCD)

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If Not RS.EOF Then
        fncGetSYNMByCode = Trim$(CStr(RS("VASYKJ")))
    End If
    RS.Close

    'BVAP01に存在しない場合はSTSP01を参照
    If fncGetSYNMByCode = "" Then
        strSQL = ""
        strSQL = strSQL & "SELECT TSSYKJ "
        strSQL = strSQL & "  FROM LIBSMF17.STSP01 "
        strSQL = strSQL & " WHERE TSDELT='' "
        strSQL = strSQL & "   AND TSSYCD=" & Val(iSYCD)

        RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
        If Not RS.EOF Then
            fncGetSYNMByCode = Trim$(CStr(RS("TSSYKJ")))
        End If
    End If

    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Function

Private Function fncGetSYCDByName(ByVal iSYNM As String) As String
    Dim CN          As Object
    Dim RS          As Object
    Dim strSQL      As String

    fncGetSYCDByName = ""
    If Trim$(iSYNM) = "" Then Exit Function

    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")

    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString

    strSQL = ""
    strSQL = strSQL & "SELECT TSSYCD "
    strSQL = strSQL & "  FROM LIBSMF17.STSP01 "
    strSQL = strSQL & " WHERE TSDELT='' "
    strSQL = strSQL & "   AND TSSYKJ='" & Replace(iSYNM, "'", "''") & "' "
    strSQL = strSQL & " FETCH FIRST 1 ROW ONLY"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If Not RS.EOF Then
        fncGetSYCDByName = Format(RS("TSSYCD"), "00000000")
    End If

    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Function

Private Sub subUpdateSheet()
    Me.Cells(1, 14) = ""
    Me.Cells(9, 2).Select
End Sub