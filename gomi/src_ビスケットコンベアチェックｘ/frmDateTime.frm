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

Option Explicit

Private Sub UserForm_Initialize()
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
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    If lblDate.Caption = "" Or lblDate.Tag = "" Then Exit Sub
    If lblTime.Caption = "" Or lblTime.Tag = "" Then Exit Sub
    If Not IsDate(lblTime.Tag) Then MsgBox ("時刻が不正です"): Exit Sub
    If fncChkExist(CDate(lblDate.Tag & " " & lblTime.Tag)) Then MsgBox ("この時刻のデータはすでに存在しています"): Exit Sub
    P_DateTime = CDate(lblDate.Tag & " " & lblTime.Tag)
    P_DateTime_FLG = True
    Unload Me
End Sub

Private Function fncChkExist(ByVal dateTime As Date) As Boolean
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
End Function

Private Sub cmdDelete_Click()
    P_DateTime_FLG = True
    P_DateTimeMode = 3      '削除
    Unload Me
End Sub

Private Sub cmdCalendar_Click()
    If Not lblDate.Tag = "" Then P_DATEC = CDate(lblDate.Tag)
    Call subOpenCalendar
    If Not P_Calendar_FLG Then Exit Sub
    lblDate.Caption = Format(P_DATEC, "m/d")
    lblDate.Tag = Format(P_DATEC, "yyyy/mm/dd")
End Sub

Private Sub subOpenCalendar()
    Dim obj As New frmCalendar
    obj.Show
    Set obj = Nothing
End Sub

Private Sub cmdBack_Click()
    If lblDate.Caption = "" Then Exit Sub
    lblDate.Tag = Format(DateAdd("d", -1, CDate(lblDate.Tag)), "yyyy/mm/dd")
    lblDate.Caption = Format(CDate(lblDate.Tag), "m/d")
End Sub

Private Sub cmdNext_Click()
    If lblDate.Caption = "" Then Exit Sub
    lblDate.Tag = Format(DateAdd("d", 1, CDate(lblDate.Tag)), "yyyy/mm/dd")
    lblDate.Caption = Format(CDate(lblDate.Tag), "m/d")
End Sub

Private Sub subKeyClick(ByVal key As String)
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
End Sub

Private Sub btn0_Click()
    Call subKeyClick("0")
End Sub

Private Sub btn0_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("0")
End Sub

Private Sub btn1_Click()
    Call subKeyClick("1")
End Sub
Private Sub btn1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subKeyClick("1")
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

