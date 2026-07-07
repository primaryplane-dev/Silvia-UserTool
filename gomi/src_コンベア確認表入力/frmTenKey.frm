VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTenKey 
   Caption         =   "数値パッド"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   OleObjectBlob   =   "frmTenKey.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTenKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/**
' * @file frmTenKey.frm
' * @brief 数値パッドフォーム (UI 層)
' * @note ユーザー数値入力の受け付けと、値の返却のみを行う
' */
Option Explicit

Private ClearFLG    As Boolean


'/**
' * @brief OKボタン押下時の処理
' */
Private Sub btnOk_Click()
    On Error GoTo ErrorHandler

    Dim strMSG      As String
    
    '入力チェック
    If Not ClearFLG Then
        If lblInputValue.Caption = "" Then Exit Sub
    End If
    strMSG = CheckKeta(lblInputValue.Caption)
    If strMSG <> "" Then
        Call MsgBox(strMSG, vbExclamation, Bas_Configuration.SYSTEM_NAME)
        Exit Sub
    End If
    P_TenKeyData = lblInputValue.Caption
    P_TenKey_FLG = True
    Unload Me

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmTenKey.btnOk_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("OKボタン処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 入力値の桁数チェック
' * @param i_Val 入力値
' * @return String エラーメッセージ（正常時は空文字）
' */
Private Function CheckKeta(ByVal i_Val As String) As String
    On Error GoTo ErrorHandler

    CheckKeta = ""
    If i_Val = "" Then Exit Function
    
    'マイナスは桁に入れない
    Dim strWK   As String:  strWK = Replace(i_Val, "-", "")
    If P_TenKeyPointMode = 0 Then
        '正の整数のみの場合
        If Len(strWK) > P_TenKeyKeta Then CheckKeta = P_TenKeyKeta & "桁以内で入力してください"
        Exit Function
    End If
    
    Dim varWK   As Variant: varWK = Split(strWK, ".")
    '整数部の桁チェック/小数部は１桁のみ(今のところ)
    If Len(varWK(0)) > P_TenKeyKeta Then CheckKeta = "整数部は" & P_TenKeyKeta & "桁以内で入力してください"
    If UBound(varWK) >= 1 Then
        If P_TenKeyPointMode = 1 Then
            If Len(varWK(1)) > 1 Then CheckKeta = "小数部は１桁以内で入力してください"
        ElseIf P_TenKeyPointMode = 2 Then
            If Len(varWK(1)) > 2 Then CheckKeta = "小数部は２桁以内で入力してください"
        End If
    End If

    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmTenKey.CheckKeta: " & Err.Number & " - " & Err.Description
    CheckKeta = "桁数チェック処理でエラーが発生しました。"
End Function


'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    lblInputValue.Caption = P_TenKeyData
    P_TenKey_FLG = False
    P_TEdit_FLG = False
    ClearFLG = False
    
    'マイナス、小数点ボタンの表示切替
    If Not P_TenKeyPointMode = 0 And P_TenKeyMinus Then
        btnD.Visible = True
        btnMinus.Visible = True
    ElseIf Not P_TenKeyPointMode = 0 And Not P_TenKeyMinus Then
        btnMinus.Visible = False
        btnD.Visible = True
        btnC.Width = 78
        btnD.Width = 78:   btnD.Left = 84
        btn0.Width = 78:   btn0.Left = 162
    Else
        btnMinus.Visible = False
        btnD.Visible = False
        btnC.Width = 78
        btn0.Width = 156:   btn0.Left = 84
    End If
    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmTenKey.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("初期化処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief キャンセルボタン押下時の処理
' */
Private Sub btnCancel_Click()
    On Error GoTo ErrorHandler

    Unload Me
    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmTenKey.btnCancel_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("キャンセル処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief テンキー入力処理
' * @param key 入力キー
' */
Private Sub subTenKeyClick(ByVal key As String)
    On Error GoTo ErrorHandler

    Dim strLabel As String
    If Not P_TEdit_FLG Then
        strLabel = ""
        P_TEdit_FLG = True
    Else
        strLabel = lblInputValue.Caption
    End If
    Select Case key
        Case "C"
            strLabel = "":  ClearFLG = True
        Case "-"
            If InStr(strLabel, "-") > 0 Then
                strLabel = Replace(strLabel, "-", "")
            Else
                strLabel = "-" & strLabel
            End If
        Case "."
            Select Case strLabel
                Case ""
                    strLabel = "0."
                Case Else
                    strLabel = strLabel & "."
            End Select
        Case Else
            strLabel = strLabel & key
    End Select
    lblInputValue.Caption = strLabel
    Exit Sub
    
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmTenKey.subTenKeyClick: " & Err.Number & " - " & Err.Description
    Call MsgBox("テンキー入力処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'------------------------------------------------------------------------------------------------
' 数値キー
'------------------------------------------------------------------------------------------------
Private Sub btnD_Click()
    Call subTenKeyClick(".")
End Sub

Private Sub btnMinus_Click()
    Call subTenKeyClick("-")
End Sub

Private Sub btnC_Click()
    Call subTenKeyClick("C")
End Sub

Private Sub btn0_Click()
    Call subTenKeyClick("0")
End Sub

Private Sub btn0_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("0")
End Sub

Private Sub btn1_Click()
    Call subTenKeyClick("1")
End Sub
Private Sub btn1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("1")
End Sub

Private Sub btn2_Click()
    Call subTenKeyClick("2")
End Sub
Private Sub btn2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("2")
End Sub

Private Sub btn3_Click()
    Call subTenKeyClick("3")
End Sub
Private Sub btn3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("3")
End Sub

Private Sub btn4_Click()
    Call subTenKeyClick("4")
End Sub
Private Sub btn4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("4")
End Sub

Private Sub btn5_Click()
    Call subTenKeyClick("5")
End Sub
Private Sub btn5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("5")
End Sub

Private Sub btn6_Click()
    Call subTenKeyClick("6")
End Sub
Private Sub btn6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("6")
End Sub

Private Sub btn7_Click()
    Call subTenKeyClick("7")
End Sub
Private Sub btn7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("7")
End Sub

Private Sub btn8_Click()
    Call subTenKeyClick("8")
End Sub
Private Sub btn8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("8")
End Sub

Private Sub btn9_Click()
    Call subTenKeyClick("9")
End Sub
Private Sub btn9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("9")
End Sub


