Attribute VB_Name = "Bas_Common"
Option Explicit

' AS400äiî[éûÇÃÉoÉCÉgêîÇãÅÇﬂÇÈ
Public Function LenAS(ByVal i_ï∂éöóÒ As String) As Integer
    Dim i       As Integer
    Dim moji_z  As Integer
    Dim ret     As Integer

    For i = 0 To (Len(i_ï∂éöóÒ) - 1)
        Select Case LenB(StrConv(Mid(i_ï∂éöóÒ, i + 1, 1), vbFromUnicode))
        Case 1
            If moji_z = 2 Then ret = ret + 1
            ret = ret + 1
            moji_z = 1
        Case 2
            If moji_z <> 2 Then ret = ret + 1
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
    Call subAllVisible
    fncGetMaxRow = stList.Cells(stList.Rows.Count, 2).End(xlUp).Row
End Function

Public Function fncChkstr(InputStr As String) As Boolean
    Dim objReg As Object
    Set objReg = CreateObject("VBScript.RegExp")
    objReg.IgnoreCase = True
    objReg.Pattern = "^[a-zA-Z0-9]+$"
    fncChkstr = objReg.Test(InputStr)
    Set objReg = Nothing
End Function

' -------------------------------------------------------------------------
' çHíˆÉRÅ[Éhïœä∑ (BFKTCD)
' éwíËÅFê¨å^=1, ó‚ãp=21
' -------------------------------------------------------------------------
Public Function fncGetKTNM(ByVal KTCD As String) As String
    fncGetKTNM = ""
    Select Case KTCD
    Case "1": fncGetKTNM = "ê¨å^"
    Case "21": fncGetKTNM = "ó‚ãp"
    Case Else: fncGetKTNM = KTCD
    End Select
End Function

Public Function fncGetKTCD(ByVal KTNM As String) As String
    fncGetKTCD = ""
    Select Case KTNM
    Case "ê¨å^": fncGetKTCD = "1"
    Case "ó‚ãp": fncGetKTCD = "21"
    Case Else: fncGetKTCD = KTNM
    End Select
End Function

Public Sub subAllVisible()
    stList.Cells.EntireRow.Hidden = False
End Sub


