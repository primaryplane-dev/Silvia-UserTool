VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "製品検索"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13035
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**
' * @file frmSearch.frm
' * @brief 製品検索画面 (UI 層)
' * @note 製品名・カナ・コードでの検索と選択を担当
' */

Option Explicit

'/**
' * @brief キャンセルボタン押下時の処理
' * @note フォームを閉じるのみ
' */
Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSearch.cmdCancel_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("キャンセル処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 登録ボタン押下時の処理
' * @note 選択データをグローバル変数に格納しフォームを閉じる
' */
Private Sub cmdRegist_Click()
    On Error GoTo ErrorHandler

    If lstData.Text = "" Then Exit Sub
    P_CalendarSelected = True
    P_SelHNO = lstData.Value
    P_SelHNM = lstData.Text
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSearch.cmdRegist_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("登録処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief フォーム初期化処理
' * @note 検索ワード・リスト初期化
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    P_CalendarSelected = False
    txtWord.Text = ""
    lstData.Clear

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSearch.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("画面初期化中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 検索ボタン押下時の処理
' * @note 検索ワードをトリムし、リストを再構築
' */
Private Sub cmdSearch_Click()
    On Error GoTo ErrorHandler

    txtWord.Text = Trim(txtWord.Text)
    If txtWord.Text = "" Then Exit Sub
    Call subEditDataList

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSearch.cmdSearch_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("検索処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 検索結果リスト編集処理
' * @note DBから検索結果を取得しリストに反映
' */
Private Sub subEditDataList()
    On Error GoTo ErrorHandler

    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String

    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString

    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT RHHNO, HMHNM"
    strSQL = strSQL & "   FROM LIBSMF17.SRHP01 "
    strSQL = strSQL & "   LEFT JOIN LIBWMF.WHMP01 "
    strSQL = strSQL & "     ON HMDLT = '' "
    strSQL = strSQL & "    AND HMHNO = RHHNO "
    strSQL = strSQL & "  WHERE (" & fncEditWhere(txtWord.Text) & ")"
    strSQL = strSQL & "    AND RHDLT = '' "

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    lstData.Clear
    Do While Not RS.EOF
        lstData.AddItem
        lstData.List(lstData.ListCount - 1, 0) = RS("RHHNO")
        lstData.List(lstData.ListCount - 1, 1) = RS("HMHNM")
        RS.MoveNext
    Loop

    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSearch.subEditDataList: " & Err.Number & " - " & Err.Description
    Call MsgBox("リスト編集処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief WHERE句生成処理
' * @param Word 検索ワード
' * @return String WHERE句
' * @note 検索ワードを全角・半角変換し部分一致条件を生成
' */
Private Function fncEditWhere(ByVal Word As String) As String
    On Error GoTo ErrorHandler

    Dim sSp()       As String: sSp = Split(Replace(Word, "　", " "), " ")
    Dim sSpZ()      As String: ReDim sSpZ(UBound(sSp))
    Dim sSpH()      As String: ReDim sSpH(UBound(sSp))
    Dim i           As Long
    Dim str         As String

    For i = 0 To UBound(sSp)
        sSpZ(i) = StrConv(sSp(i), vbWide)
        sSpH(i) = StrConv(sSp(i), vbNarrow)
    Next
    fncEditWhere = ""
    For i = 0 To UBound(sSp)
        If Not fncEditWhere = "" Then fncEditWhere = fncEditWhere & " AND "
        str = ""
        str = str & "    HMKNM LIKE '%" & sSp(i) & "%' OR HMHNM LIKE '%" & sSp(i) & "%'"
        str = str & " OR HMKNM LIKE '%" & sSpZ(i) & "%' OR HMHNM LIKE '%" & sSpZ(i) & "%'"
        str = str & " OR HMKNM LIKE '%" & sSpH(i) & "%' OR HMHNM LIKE '%" & sSpH(i) & "%'"
        fncEditWhere = fncEditWhere & "(" & str & ")"
    Next
    Exit Function
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSearch.fncEditWhere: " & Err.Number & " - " & Err.Description
    Call MsgBox("WHERE句生成処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
    fncEditWhere = ""
End Function
