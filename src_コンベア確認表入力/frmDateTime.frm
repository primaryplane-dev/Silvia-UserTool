VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDateTime 
   Caption         =   "時刻入力"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   OleObjectBlob   =   "frmDateTime.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
/**
 * @file frmDateTime.frm
 * @brief 日付・時刻入力画面 (UI 層)
 * @note ユーザーによる日付・時刻入力と返却のみを担当
 */

Option Explicit


'/**
' * @brief フォーム初期化処理
' */
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    P_DateTime_FLG = False
    If Not P_DateTimeMode = 2 Then  '編集モード
        cmdDelete.Enabled = False
    End If
    If Not P_DateTime = 0 Then
        lblDate.Caption = Format(P_DateTime, "m/d")
        lblDate.Tag = Format(P_DateTime, "yyyy/mm/dd")
        lblTime.Caption = Format(P_DateTime, "hh:mm")
        lblTime.Tag = Format(P_DateTime, "hh:mm:ss")
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.UserForm_Initialize: " & Err.Number & " - " & Err.Description
    Call MsgBox("画面初期化中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief キャンセルボタン押下時の処理
' */
Private Sub btnCancel_Click()
    On Error GoTo ErrorHandler
    Unload Me
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.btnCancel_Click: " & Err.Number & " - " & Err.Description
	Call MsgBox("キャンセル処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 決定ボタン押下時の処理
' */
Private Sub btnOk_Click()
    On Error GoTo ErrorHandler
    If lblDate.Caption = "" Or lblDate.Tag = "" Then Exit Sub
    If lblTime.Caption = "" Or lblTime.Tag = "" Then Exit Sub
    If Not IsDate(lblTime.Tag) Then
        Call MsgBox("時刻が不正です", vbExclamation, Bas_Configuration.SYSTEM_NAME)
        Exit Sub
    End If
    If fncChkExist(CDate(lblDate.Tag & " " & lblTime.Tag)) Then
        Call MsgBox("この時刻のデータはすでに存在しています", vbExclamation, Bas_Configuration.SYSTEM_NAME)
        Exit Sub
    End If
    P_DateTime = CDate(lblDate.Tag & " " & lblTime.Tag)
    P_DateTime_FLG = True
    Unload Me
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.btnOk_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("決定処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 入力済み時刻の重複チェック
' */
Private Function fncChkExist(ByVal dateTime As Date) As Boolean
    On Error GoTo ErrorHandler
    Dim lRow        As Long
    fncChkExist = False
    lRow = 1
    Do While Not stWork.Cells(lRow, 1) = ""
        If Not stWork.Cells(lRow, 14) = "X" And stWork.Cells(lRow, 1) = Format(dateTime, "yyyymmddhhmmss") Then
            fncChkExist = True
            Exit Do
        End If
        lRow = lRow + 1
    Loop
    Exit Function
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.fncChkExist: " & Err.Number & " - " & Err.Description
    Call MsgBox("時刻重複チェック中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
    fncChkExist = False
End Function

'/**
' * @brief 削除ボタン押下時の処理
' */
Private Sub cmdDelete_Click()
    On Error GoTo ErrorHandler
    P_DateTime_FLG = True
    P_DateTimeMode = 3      '削除
    Unload Me
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.cmdDelete_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("削除処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief カレンダーボタン押下時の処理
' */
Private Sub cmdCalendar_Click()
    On Error GoTo ErrorHandler
    If Not lblDate.Tag = "" Then P_DATEC = CDate(lblDate.Tag)
    Call subOpenCalendar
    If Not P_Calendar_FLG Then Exit Sub
    lblDate.Caption = Format(P_DATEC, "m/d")
    lblDate.Tag = Format(P_DATEC, "yyyy/mm/dd")
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.cmdCalendar_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("カレンダー処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief カレンダー画面の表示
' */
Private Sub subOpenCalendar()
    On Error GoTo ErrorHandler
    Dim obj As New frmCalendar
    obj.Show
    Set obj = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.subOpenCalendar: " & Err.Number & " - " & Err.Description
    Call MsgBox("カレンダー表示中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 前日ボタン押下時の処理
' */
Private Sub cmdBack_Click()
    On Error GoTo ErrorHandler
    If lblDate.Caption = "" Then Exit Sub
    lblDate.Tag = Format(DateAdd("d", -1, CDate(lblDate.Tag)), "yyyy/mm/dd")
    lblDate.Caption = Format(CDate(lblDate.Tag), "m/d")
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.cmdBack_Click: " & Err.Number & " - " & Err.Description
	Call MsgBox("前日処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 翌日ボタン押下時の処理
' */
Private Sub cmdNext_Click()
    On Error GoTo ErrorHandler
    If lblDate.Caption = "" Then Exit Sub
    lblDate.Tag = Format(DateAdd("d", 1, CDate(lblDate.Tag)), "yyyy/mm/dd")
    lblDate.Caption = Format(CDate(lblDate.Tag), "m/d")
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.cmdNext_Click: " & Err.Number & " - " & Err.Description
	Call MsgBox("翌日処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief テンキー入力処理
' */
Private Sub subKeyClick(ByVal key As String)
    On Error GoTo ErrorHandler
    Dim strLabel As String
    strLabel = lblTime.Caption
    Select Case key
        Case "C"
            strLabel = ""
        Case ":"
            If InStr(1, strLabel, ":") = 0 Then
                strLabel = strLabel & ":"
            End If
        Case Else
            strLabel = strLabel & key
    End Select
    lblTime.Caption = strLabel
    lblTime.Tag = strLabel & IIf(strLabel = "", "", ":00")
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] frmDateTime.subKeyClick: " & Err.Number & " - " & Err.Description
    Call MsgBox("テンキー入力処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

Private Sub btn2_Click()
    Call subKeyClick("2")
End Sub
Private Sub btn2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("2")
End Sub

Private Sub btn3_Click()
    Call subKeyClick("3")
End Sub
Private Sub btn3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("3")
End Sub

Private Sub btn4_Click()
    Call subKeyClick("4")
End Sub
Private Sub btn4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("4")
End Sub

Private Sub btn5_Click()
    Call subKeyClick("5")
End Sub
Private Sub btn5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("5")
End Sub

Private Sub btn6_Click()
    Call subKeyClick("6")
End Sub
Private Sub btn6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("6")
End Sub

Private Sub btn7_Click()
    Call subKeyClick("7")
End Sub
Private Sub btn7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("7")
End Sub

Private Sub btn8_Click()
    Call subKeyClick("8")
End Sub
Private Sub btn8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("8")
End Sub

Private Sub btn9_Click()
    Call subKeyClick("9")
End Sub
Private Sub btn9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("9")
End Sub

Private Sub btnC_Click()
    Call subKeyClick("C")
End Sub

Private Sub btnColon_Click()
    Call subKeyClick(":")
End Sub

Private Sub btnNow_Click()
    Dim dateWk      As Date
    Dim strWK       As String
'    strWk = Format(Now, "yyyymmddhhmmss")
'    dateWk = CDate(Left(strWk, 4) & "/" & Mid(strWk, 5, 2) & "/" & Mid(strWk, 7, 2) & " " & Mid(strWk, 9, 2) & ":" & Mid(strWk, 11, 2) & ":" & Right(strWk, 2))
    dateWk = Now
    lblDate.Tag = Format(dateWk, "yyyy/mm/dd")
    lblDate.Caption = Format(dateWk, "m/d")
    lblTime.Tag = Format(dateWk, "hh:mm:00")
    lblTime.Caption = Format(dateWk, "hh:mm")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'フォームの「Ｘ」ボタンは使用不可にする
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

