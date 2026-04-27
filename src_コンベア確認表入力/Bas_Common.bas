Attribute VB_Name = "Bas_Common"
Option Explicit

' AS400格納時のバイト数を求める(半角 1 バイト、全角 2 バイト)
'  ２バイト文字の前後に制御文字が入ることを考慮する
Public Function LenAS(ByVal i_文字列 As String) As Integer
    Dim i       As Integer
    Dim moji_z  As Integer     '前回の文字のバイト数
    Dim ret     As Integer     '長さ

    For i = 0 To (Len(i_文字列) - 1)
        Select Case LenB(StrConv(Mid(i_文字列, i + 1, 1), vbFromUnicode))
        Case 1
            If moji_z = 2 Then ret = ret + 1   '制御文字(2byte文字の終了)
            ret = ret + 1
            moji_z = 1
        Case 2
            If moji_z <> 2 Then ret = ret + 1  '制御文字(2byte文字の開始)
            ret = ret + 2
            moji_z = 2
        End Select
    Next
    If moji_z = 2 Then ret = ret + 1

    LenAS = ret
End Function

Public Sub subBeforeEdit()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Public Sub subAfterEdit()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Function fncGetMaxRow() As Long
    fncGetMaxRow = stList.Cells(stList.Rows.Count, 6).End(xlUp).Row
End Function

Public Function fncGetKTNM(ByVal KTCD As String) As String
    fncGetKTNM = ""
    Select Case KTCD
    Case "1": fncGetKTNM = "成型"
    Case "21": fncGetKTNM = "冷却"
    End Select
End Function
