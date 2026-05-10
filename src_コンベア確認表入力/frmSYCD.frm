VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSYCD 
   Caption         =   "担当者"
   ClientHeight    =   3270
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
' * @brief 担当者選択フォーム (UI 層)
' * @note ユーザー入力の受け付けと、担当者情報の取得・返却のみを行う
' */
Option Explicit


'/**
' * @brief 確定ボタン押下時の処理
' */
Private Sub cmdOK_Click()
    On Error GoTo ErrorHandler

    Dim strSYCD     As String
    Dim strSYNM     As String
    
    '条件を保存する(ComboBox優先)
    If cbSYCD.Text <> "" Then
        strSYCD = Format(cbSYCD.Value, "00000000")
        strSYNM = cbSYCD.Text
    ElseIf lblSYNM.Caption <> "" Then
        strSYCD = Format(txtSYCD.Text, "00000000")
        strSYNM = lblSYNM.Caption
    End If
    If strSYCD = "" Then Exit Sub
    P_SYCD = strSYCD
    P_SYNM = strSYNM
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.cmdOK_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("確定処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 数値パッドボタン押下時の処理
' */
Private Sub cmdInputSYCD_Click()
    On Error GoTo ErrorHandler

    P_TenKeyData = txtSYCD.Text
    P_TenKeyKeta = 8:   P_TenKeyMinus = False:  P_TenKeyPointMode = 0
    frmTenKey.Show
    
    If P_TenKey_FLG Then
        txtSYCD.Text = P_TenKeyData
        lblSYNM.Caption = ""
        Call GetSYNM
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.cmdInputSYCD_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("数値パッド処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 社員マスタの存在チェック＆社員名取得
' */
Private Sub GetSYNM()
    On Error GoTo ErrorHandler

    Dim strSYCD     As String:  strSYCD = txtSYCD.Text
    Dim strSYNM     As String
    
    '入力チェック
    If Val(strSYCD) = 0 Then Exit Sub
    
    '社員マスタを読む
    If Not GetSYCD(strSYCD, strSYNM) Then
        Call MsgBox("社員コードが間違っています", vbExclamation, Bas_Configuration.SYSTEM_NAME)
        Exit Sub
    End If
    txtSYCD.Text = strSYCD
    lblSYNM.Caption = strSYNM
    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.GetSYNM: " & Err.Number & " - " & Err.Description
    Call MsgBox("社員名取得処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Call MakecbSYCD
    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("初期化処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 担当者コンボボックス生成処理
' */
Public Sub MakecbSYCD()
    On Error GoTo ErrorHandler

    Dim i           As Integer
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    'STSP01:   担当者マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT "
    strSQL = strSQL & " TSSYCD, TSSYKJ "
    strSQL = strSQL & " FROM  LIBSMF17.STSP01  "
    strSQL = strSQL & " WHERE TSDELT='' "
    strSQL = strSQL & " ORDER BY TSSYCD"
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    cbSYCD.Clear
    cbSYCD.AddItem
    Do While Not RS.EOF
        cbSYCD.AddItem
        cbSYCD.List(cbSYCD.ListCount - 1, 0) = Trim(RS("TSSYCD"))
        cbSYCD.List(cbSYCD.ListCount - 1, 1) = Trim(RS("TSSYKJ"))
        RS.MoveNext
    Loop
    
    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.MakecbSYCD: " & Err.Number & " - " & Err.Description
    Call MsgBox("担当者リスト生成処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 社員コードから社員名を取得
' * @param iSYCD 入力社員コード
' * @param oSYNM 取得社員名（ByRef）
' * @return Boolean 取得成功時True
' */
Private Function GetSYCD(ByVal iSYCD As String, ByRef oSYNM As String) As Boolean
    On Error GoTo ErrorHandler

    Dim CN          As Object:  Set CN = CreateObject("ADODB.Connection")
    Dim RS          As Object:  Set RS = CreateObject("ADODB.Recordset")
    Dim strSQL      As String
    GetSYCD = False
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
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
    Exit Function
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSYCD.GetSYCD: " & Err.Number & " - " & Err.Description
    Call MsgBox("社員名取得処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
    GetSYCD = False
End Function


