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

'/**
' * @file frmSelect.frm
' * @brief 条件指定フォーム (UI 層)
' * @note ユーザー入力の受け付けと、ロジック層の呼び出し、結果の受け渡しのみを行う
' */
Option Explicit


'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    'Call subMakeCombo
    If Not P_DATE = 0 Then txtDate.Text = Format(P_DATE, "yyyy/mm/dd")
    If Not P_HINO = "" Then txtHINM.Tag = P_HINO: txtHINM.Text = P_HINM
    If Not P_KJNO = "" Then cmbKJNM.Value = P_KJNO & "@" & P_KJNM
    If Not P_SHBU = "" Then cmbSHBU.Value = P_SHBU
    If P_ChkKBN = 1 Then
        optKT1.Value = True
    ElseIf P_ChkKBN = 2 Then
        optKT2.Value = True
    End If
    P_Regist = False

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("初期化処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief キャンセルボタン押下時の処理
' */
Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.cmdCancel_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("キャンセル処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 登録ボタン押下時の処理
' */
Private Sub cmdRegist_Click()
    On Error GoTo ErrorHandler

    If txtDate.Text = "" Then Call MsgBox("生産日を選択してください", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    If txtHINM.Text = "" Or txtHINM.Tag = "" Then Call MsgBox("品名を選択してください", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    If cmbKJNM.Enabled Then
        If cmbKJNM.Text = "" Or cmbKJNM.Value = "" Then Call MsgBox("生地名を選択してください", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    End If
    If CDate(txtDate.Text) > Date Then Call MsgBox("未来日は選択できません。", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    'If cmbSHBU.Text = "" Then Call MsgBox("アイテムを選択してください", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    If Not optKT1.Value And Not optKT2.Value Then Call MsgBox("チェック区分を選択してください", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    
    Dim sSp() As String
    P_DATE = CDate(txtDate.Text)
    P_HINO = txtHINM.Tag
    P_HINM = txtHINM.Text
    P_KJNO = ""
    P_KJNM = ""
    'P_SHBU = cmbSHBU.Value
    P_SHNM = cmbSHBU.Text
    P_ChkKBN = IIf(optKT1.Value, 1, IIf(optKT2.Value, 2, 0))
    'P_FNCD = False
    If cmbKJNM.Enabled Then
        sSp = Split(cmbKJNM.Value, "@")
        P_KJNO = sSp(0)
        P_KJNM = sSp(1)
    End If
    P_Regist = True
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.cmdRegist_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("登録処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
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


'/**
' * @brief カレンダーボタン押下時の処理
' */
Private Sub cmdCalendar_Click()
    On Error GoTo ErrorHandler

    If Not txtDate.Text = "" Then P_DATEC = CDate(txtDate.Text)
    Call subOpenCalendar
    If Not P_Calendar_FLG Then Exit Sub
    txtDate.Text = Format(P_DATEC, "yyyy/mm/dd")

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.cmdCalendar_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("カレンダー表示処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief カレンダーフォームを開く
' */
Private Sub subOpenCalendar()
    On Error GoTo ErrorHandler

    Dim obj As New frmCalendar
    obj.Show
    Set obj = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.subOpenCalendar: " & Err.Number & " - " & Err.Description
    Call MsgBox("カレンダー呼び出しでエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 品名テキスト変更時の処理
' */
Private Sub txtHINM_Change()
    On Error GoTo ErrorHandler

    cmbKJNM.Clear
    If Not txtHINM.Text = "" Then
        Call subMakeKJCombo
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.txtHINM_Change: " & Err.Number & " - " & Err.Description
    Call MsgBox("品名変更時の処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 検索ボタン押下時の処理
' */
Private Sub cmdSerch_Click()
    On Error GoTo ErrorHandler

    Call subOpenSerch
    If Not P_CalendarSelected Then Exit Sub
    txtHINM.Tag = P_SelHNO
    txtHINM.Text = P_SelHNM

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.cmdSerch_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("検索処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 検索フォームを開く
' */
Private Sub subOpenSerch()
    On Error GoTo ErrorHandler

    Dim obj As New frmSearch
    obj.Show
    Set obj = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.subOpenSerch: " & Err.Number & " - " & Err.Description
    Call MsgBox("検索フォーム呼び出しでエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 生地名コンボボックスを生成
' */
Private Sub subMakeKJCombo()
    On Error GoTo ErrorHandler

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

    Exit Sub
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelect.subMakeKJCombo: " & Err.Number & " - " & Err.Description
    Call MsgBox("生地名コンボ生成でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub
