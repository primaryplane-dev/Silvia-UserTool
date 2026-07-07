VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectSHBU 
   Caption         =   "条件選択"
   ClientHeight    =   3336
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6735
   OleObjectBlob   =   "frmSelectSHBU.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectSHBU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**
' * @file frmSelectSHBU.frm
' * @brief 条件選択画面 (UI 層)
' * @note ユーザーによる支部選択と、選択結果の返却のみを担当
' */

Option Explicit

'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    P_CalendarSelected = False

    Call subMakeCombo

    If Not P_SHBU = "" Then
        cmbSHBU.Value = P_SHBU
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectSHBU.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("画面初期化中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 支部コンボボックスの生成処理
' */
Private Sub subMakeCombo()
    On Error GoTo ErrorHandler

    Dim arrItems As Variant
    arrItems = Bas_LogicConveyor.GetSHBUList()
    Bas_Utilities.FillComboBox cmbSHBU, arrItems

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectSHBU.subMakeCombo: " & Err.Number & " - " & Err.Description
    Call MsgBox("支部リスト生成中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 登録ボタン押下時の処理
' */
Private Sub cmdRegist_Click()
    On Error GoTo ErrorHandler

    If cmbSHBU.Text = "" Then Exit Sub
    P_SHBU = cmbSHBU.Value
    P_SHNM = cmbSHBU.Text
    P_CalendarSelected = True
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectSHBU.cmdRegist_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("登録処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief キャンセルボタン押下時の処理
' */
Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    Unload Me

    Exit Sub
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmSelectSHBU.cmdCancel_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("キャンセル処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

