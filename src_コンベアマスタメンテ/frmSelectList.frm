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
' * @brief 選択リスト画面 (UI 層)
' * @note ユーザーによる項目選択と、選択結果の返却のみを担当
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
    Call MsgBox("キャンセル処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief リスト項目クリック時の処理
' */
Private Sub lstItem_Click()
    On Error GoTo ErrorHandler

    If lstItem.ListIndex < 0 Then Exit Sub
    P_SelItem = lstItem.List(lstItem.ListIndex)
    P_Regist = True
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectList.lstItem_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("選択処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Call subMakeList
    P_SelItem = ""
    P_Regist = False

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectList.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("画面初期化中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 選択リストの生成処理
' */
Private Sub subMakeList()
    On Error GoTo ErrorHandler

    Dim arrItems As Variant
    arrItems = Bas_LogicConveyor.GetSelectableProcessList()
    Bas_Utilities.FillListBox lstItem, arrItems

    Exit Sub
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectList.subMakeList: " & Err.Number & " - " & Err.Description
    Call MsgBox("リスト生成中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub
