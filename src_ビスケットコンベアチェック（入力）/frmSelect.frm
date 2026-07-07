VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelect 
   Caption         =   "条件指定"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   OleObjectBlob   =   "frmSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub UserForm_Initialize()
''    Call subMakeCombo
'    If Not P_DATE = 0 Then txtDate.Text = Format(P_DATE, "yyyy/mm/dd")
'    If Not P_HINO = "" Then txtHINM.Tag = P_HINO: txtHINM.Text = P_HINM
'    If Not P_KJNO = "" Then cmbKJNM.Value = P_KJNO & "@" & P_KJNM
'    If Not P_SHBU = "" Then cmbSHBU.Value = P_SHBU
'    If P_ChkKBN = 1 Then
'        optKT1.Value = True
'    ElseIf P_ChkKBN = 2 Then
'        optKT2.Value = True
'    End If
'    P_Regist = False
'End Sub

Private Sub UserForm_Initialize()
'    Call subMakeCombo
    If Not P_DATE = 0 Then txtDate.Text = Format(P_DATE, "yyyy/mm/dd")
    If Not P_HINO = "" Then txtHINM.Tag = P_HINO: txtHINM.Text = P_HINM
    If P_KJNO <> "" Then
        Dim targetKJ As String
        Dim i As Integer
        targetKJ = P_KJNO & "@" & P_KJNM
        For i = 0 To cmbKJNM.ListCount - 1
            If cmbKJNM.List(i, 0) = targetKJ Then
                cmbKJNM.ListIndex = i
                Exit For
            End If
        Next
    End If
    If P_SHBU <> "" Then cmbSHBU.Value = P_SHBU
    If P_ChkKBN = 1 Then
        optKT1.Value = True
    ElseIf P_ChkKBN = 2 Then
        optKT2.Value = True
    End If
    P_Regist = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRegist_Click()
    If txtDate.Text = "" Then MsgBox ("生産日を選択してください"): Exit Sub
    If txtHINM.Text = "" Or txtHINM.Tag = "" Then MsgBox ("品名を選択してください"): Exit Sub
    If cmbKJNM.Enabled Then
        If cmbKJNM.Text = "" Or cmbKJNM.Value = "" Then MsgBox ("生地名を選択してください"): Exit Sub
    End If
    If CDate(txtDate.Text) > Date Then MsgBox ("未来日は選択できません。"): Exit Sub
'    If cmbSHBU.Text = "" Then MsgBox ("アイテムを選択してください"): Exit Sub
    If Not optKT1.Value And Not optKT2.Value Then MsgBox ("チェック区分を選択してください"): Exit Sub
    Dim sSp()   As String
    
    P_DATE = CDate(txtDate.Text)
    P_HINO = txtHINM.Tag
    P_HINM = txtHINM.Text
    P_KJNO = ""
    P_KJNM = ""
'    P_SHBU = cmbSHBU.Value
    P_SHNM = cmbSHBU.Text
    P_ChkKBN = IIf(optKT1.Value, 1, IIf(optKT2.Value, 2, 0))
'    P_FNCD = False
    If cmbKJNM.Enabled Then
        sSp = Split(cmbKJNM.Value, "@")
        P_KJNO = sSp(0)
        P_KJNM = sSp(1)
    End If
    
    P_Regist = True
    Unload Me
End Sub

'Private Sub subMakeCombo()
'    cmbSHBU.AddItem
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 0) = ""
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 1) = ""
'    cmbSHBU.AddItem
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 0) = "1"
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 1) = "ビスケット"
'    cmbSHBU.AddItem
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 0) = "2"
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 1) = "クッキー"
'    cmbSHBU.AddItem
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 0) = "3"
'    cmbSHBU.List(cmbSHBU.ListCount - 1, 1) = "ドーナツ"
'End Sub

Private Sub cmdCalendar_Click()
    If Not txtDate.Text = "" Then P_DATEC = CDate(txtDate.Text)
    Call subOpenCalendar
    If Not P_Calendar_FLG Then Exit Sub
    txtDate.Text = Format(P_DATEC, "yyyy/mm/dd")
End Sub

Private Sub subOpenCalendar()
    Dim obj As New frmCalendar
    obj.Show
    Set obj = Nothing
End Sub

Private Sub txtHINM_Change()
    cmbKJNM.Clear
    If Not txtHINM.Text = "" Then
        Call subMakeKJCombo
    End If
End Sub

Private Sub cmdSerch_Click()
    Call subOpenSerch
    If Not P_Regist2 Then Exit Sub
    txtHINM.Tag = P_SelHNO
    txtHINM.Text = P_SelHNM
End Sub

Private Sub subOpenSerch()
    Dim obj As New frmSearch
    obj.Show
    Set obj = Nothing
End Sub

Private Sub subMakeKJCombo()
    Dim CN      As New ADODB.Connection
    Dim RS      As New ADODB.Recordset
    Dim strSQL  As String
    Dim i       As Integer
    Dim j       As Integer
    Dim arrK()  As String
    Dim bExist  As Boolean
    
    cmbKJNM.Enabled = True
    'DB接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT "
    For i = 1 To 10
        If i > 1 Then strSQL = strSQL & ","
        strSQL = strSQL & "GRG" & Format(i, "00") & ", HG" & i & ".HGHNM AS KGNM" & i
    Next
    strSQL = strSQL & "   FROM LIBNMF17.NGRP01 "
    For i = 1 To 10
        strSQL = strSQL & "   LEFT JOIN LIBNMF17.NHGP01 AS HG" & i
        strSQL = strSQL & "     ON HG" & i & ".HGDLT = '' "
        strSQL = strSQL & "    AND HG" & i & ".HGHN4 = GRG" & Format(i, "00")
    Next
    strSQL = strSQL & "  WHERE GRDLT = '' "
    strSQL = strSQL & "    AND GRHNO =" & Val(txtHINM.Tag)
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    Do While Not RS.EOF
        ReDim arrK(0)
        For i = 1 To 10
            bExist = False
            If Left(RS("GRG" & Format(i, "00")), 1) = "K" Then
                For j = 1 To UBound(arrK)
                    If arrK(j) = RS("GRG" & Format(i, "00")) & "@" & RS("KGNM" & i) Then bExist = True: Exit For
                Next
                If Not bExist Then
                    ReDim Preserve arrK(UBound(arrK) + 1)
                    arrK(UBound(arrK)) = RS("GRG" & Format(i, "00")) & "@" & RS("KGNM" & i)
                End If
            End If
        Next
        If UBound(arrK) > 1 Then
            For i = 1 To UBound(arrK)
                cmbKJNM.AddItem
                cmbKJNM.List(cmbKJNM.ListCount - 1, 0) = arrK(i)
                cmbKJNM.List(cmbKJNM.ListCount - 1, 1) = Left(arrK(i), InStr(arrK(i), "@") - 1) & IIf(Mid(arrK(i), InStr(arrK(i), "@") + 1) = "", "", " : " & Mid(arrK(i), InStr(arrK(i), "@") + 1))
            Next
        End If
        RS.MoveNext
    Loop
    If cmbKJNM.ListCount = 0 Then cmbKJNM.Enabled = False
    
    'DB切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
End Sub
