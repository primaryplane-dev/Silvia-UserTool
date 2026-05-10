VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSYCD 
   Caption         =   "担当者"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   OleObjectBlob   =   "frmSYCD.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSYCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**
' * @file frmSYCD.frm
' * @brief 担当者選択画面 (UI 層)
' * @note ユーザーによる担当者選択・社員コード入力・社員名取得を担当
' */

Option Explicit

'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Call MakecbSYCD

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("画面初期化中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 確定ボタン押下時の処理
' */
Private Sub cmdOK_Click()
    On Error GoTo ErrorHandler

    Dim strSYCD     As String
    Dim strSYNM     As String

    '条件を保存する(ComboBox優先)
    If cbSYCD.Text <> "" Then
        strSYCD = cbSYCD.Value
        strSYNM = cbSYCD.Text
    Else
        If Not txtSYCD.Text = "" Then
            If Not fncGetSYNM Then
                Call MsgBox("社員コードが間違っています", vbExclamation, Bas_Configuration.SYSTEM_NAME)
                Exit Sub
            End If
            strSYCD = txtSYCD.Text
            strSYNM = lblSYNM.Caption
        End If
    End If
    If strSYCD = "" Then Exit Sub
    P_SYCD = Format(strSYCD, "00000000")
    P_SYNM = strSYNM
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.cmdOK_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("確定処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 数値パッドボタン押下時の処理
' */
Private Sub cmdInputSYCD_Click()
    On Error GoTo ErrorHandler

    P_TenKeyData = txtSYCD.Text
    P_TenKeyKeta = 8:   P_TenKeyMinus = False:  P_TenKeyPoint = False
    Call subOpenTenkey
    If P_TenKey_FLG Then
        txtSYCD.Text = P_TenKeyData
        lblSYNM.Caption = ""
        If Not fncGetSYNM Then
            Call MsgBox("社員コードが間違っています", vbExclamation, Bas_Configuration.SYSTEM_NAME)
        End If
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.cmdInputSYCD_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("数値パッド処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 数値パッド画面の表示
' */
Private Sub subOpenTenkey()
    On Error GoTo ErrorHandler

    Dim obj As New frmTenKey
    obj.Show
    Set obj = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.subOpenTenkey: " & Err.Number & " - " & Err.Description
    Call MsgBox("数値パッド表示中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 社員マスタの存在チェック＆社員名を取得する
' */
Private Function fncGetSYNM() As Boolean
    On Error GoTo ErrorHandler

    Dim strSYCD     As String:  strSYCD = txtSYCD.Text
    Dim strSYNM     As String
    fncGetSYNM = False

    '入力チェック
    If Val(strSYCD) = 0 Then Exit Function

    '社員マスタを読む
    If Not GetSYCD(strSYCD, strSYNM) Then Exit Function
    txtSYCD.Text = strSYCD
    lblSYNM.Caption = strSYNM
    fncGetSYNM = True

    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.fncGetSYNM: " & Err.Number & " - " & Err.Description
    Call MsgBox("社員名取得処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
    fncGetSYNM = False
End Function

'/**
' * @brief 担当者コンボボックスの生成処理
' */
Public Sub MakecbSYCD()
    On Error GoTo ErrorHandler

    Dim CN          As Object
    Dim RS          As Object
    Dim strSQL      As String

    'ＤＢ接続（共通化）
    Set CN = Bas_DbConnection.GetConnection()
    CN.CursorLocation = adUseClient ' 念のため

    'IHTP01:   担当者マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " HTSYCD, HTSYKJ "
    strSQL = strSQL & " FROM  LIBIMF.IHTP01  "
    strSQL = strSQL & " WHERE HTDELT='' "
    strSQL = strSQL & "   AND HTKJCD='" & P_KJCD & "' "
    strSQL = strSQL & " ORDER BY HTSYCD"
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    cbSYCD.Clear
    cbSYCD.AddItem
    Do While Not RS.EOF
        cbSYCD.AddItem
        cbSYCD.List(cbSYCD.ListCount - 1, 0) = Trim(RS("HTSYCD"))
        cbSYCD.List(cbSYCD.ListCount - 1, 1) = Trim(RS("HTSYKJ"))
        RS.MoveNext
    Loop

    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.MakecbSYCD: " & Err.Number & " - " & Err.Description
    Call MsgBox("担当者リスト取得処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

Private Function GetSYCD(ByVal iSYCD As String, ByRef oSYNM As String) As Boolean
    Dim CN          As Object:  Set CN = Bas_DbConnection.GetConnection()
    Dim RS          As Object:  Set RS = CreateObject("ADODB.Recordset")
    Dim strSQL      As String
    
    GetSYCD = False
    'ＤＢ接続（共通化）
    CN.CursorLocation = adUseClient ' 念のため

    'BVAP01：従業員管理マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT VASYCD, VASYKJ "
    strSQL = strSQL & " FROM  LIBBMF.BVAP01  "
    strSQL = strSQL & " WHERE VAKYUK='' "
    strSQL = strSQL & "   AND VASYCD= " & Val(iSYCD)
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    Do While Not RS.EOF
        GetSYCD = True
        oSYNM = RS("VASYKJ")
        Exit Do
    Loop

    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing
End Function


