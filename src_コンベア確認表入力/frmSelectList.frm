VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectList 
   Caption         =   "選択"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   OleObjectBlob   =   "frmSelectList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'/**
' * @file frmSelectList.frm
' * @brief 選択リストフォーム (UI 層)
' * @note ユーザー選択の受け付けと、選択値の返却のみを行う
' */
Option Explicit

'/**
' * @brief キャンセルボタン押下時の処理
' */
Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectList.cmdCancel_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("キャンセル処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief リスト項目クリック時の処理
' */
Private Sub lstItem_Click()
    On Error GoTo ErrorHandler

    P_SelCD = lstItem.List(lstItem.ListIndex, 0)
    P_SelNM = lstItem.List(lstItem.ListIndex, 1)
    P_Regist3 = True
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectList.lstItem_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("リスト選択処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Call subMakeList
    P_SelCD = ""
    P_SelNM = ""
    P_Regist3 = False

    Exit Sub
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectList.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("初期化処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief リスト生成処理
' */
Private Sub subMakeList()
    On Error GoTo ErrorHandler

    Dim varKey      As Variant
    lstItem.Clear
    lstItem.AddItem
    lstItem.List(lstItem.ListCount - 1, 0) = ""
    lstItem.List(lstItem.ListCount - 1, 1) = ""
    If P_SelectKBN = 1 Then
        lstItem.AddItem
        lstItem.List(lstItem.ListCount - 1, 0) = "1"
        lstItem.List(lstItem.ListCount - 1, 1) = "〇"   '漢字の丸じゃないとoffice2016では表示がおかしくなる
        lstItem.AddItem
        lstItem.List(lstItem.ListCount - 1, 0) = "0"
        lstItem.List(lstItem.ListCount - 1, 1) = "×"
        lstItem.AddItem
        lstItem.List(lstItem.ListCount - 1, 0) = "2"
        lstItem.List(lstItem.ListCount - 1, 1) = "－"
    Else
        Dim CN      As New ADODB.Connection
        Dim RS      As New ADODB.Recordset
        Dim strSQL  As String
        'DB接続
        CN.CursorLocation = adUseClient
        CN.Open P_ConnectString
        strSQL = ""
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "   TSSYCD "
        strSQL = strSQL & "  ,TSSYKJ "
        strSQL = strSQL & "  FROM LIBSMF17.STSP01 "
        strSQL = strSQL & " WHERE TSDELT = '' "
        strSQL = strSQL & "   AND TSTTKB = '4'"
        RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
        Do While Not RS.EOF
            lstItem.AddItem
            lstItem.List(lstItem.ListCount - 1, 0) = RS("TSSYCD")
            lstItem.List(lstItem.ListCount - 1, 1) = RS("TSSYKJ")
            RS.MoveNext
        Loop
        'DB切断
        RS.Close: Set RS = Nothing
        CN.Close: Set CN = Nothing
    End If

    Exit Sub
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectList.subMakeList: " & Err.Number & " - " & Err.Description
    Call MsgBox("リスト生成処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub
