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
Option Explicit

'/**
' * @file frmSelectList.frm
' * @brief 選択リスト画面 (UI 層)
' * @note ユーザーによる項目選択と、選択結果の返却のみを担当
' */

'×ボタンで閉じた場合もHideで破棄防止
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub


Public P_Regist As Boolean
Public P_SelItem As String

'/**
' * @brief キャンセルボタン押下時の処理
' */
Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    Me.Hide
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
    Me.P_SelItem = lstItem.List(lstItem.ListIndex)
    Me.P_Regist = True
    Me.Hide

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
    Me.P_SelItem = ""
    Me.P_Regist = False

    ' リストボックスの選択可能設定を強制
    With lstItem
        .Enabled = True
        .Locked = False
        .ListStyle = fmListStylePlain ' 0
        .MultiSelect = fmMultiSelectSingle ' 0
    End With

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
