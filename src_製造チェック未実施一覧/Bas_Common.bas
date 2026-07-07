Attribute VB_Name = "Bas_Common"
Option Explicit

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

Public Function fncChkK(ByVal i_KBN As Integer) As String
    Dim CN      As ADODB.Connection
    Dim RS      As ADODB.Recordset
    Dim strSQL  As String
    
    fncChkK = ""
    'DB接続
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM LIBSMF17.STSP01 "
    strSQL = strSQL & "  WHERE TSDELT = '' "
    strSQL = strSQL & "    AND TSTTKB = '" & i_KBN & "' "
    strSQL = strSQL & "    AND TSSYCD = " & Val(P_SYCD)
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If RS.RecordCount > 0 Then
        fncChkK = RS("TSSYKJ")
    End If
    
    'DB切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
End Function

Public Function fncGetSHNM(ByVal SHBU As String) As String
    fncGetSHNM = ""
    Select Case SHBU
    Case "1": fncGetSHNM = "ビスケット"
    Case "2": fncGetSHNM = "クッキー"
    Case "3": fncGetSHNM = "ドーナツ"
    End Select
End Function

Public Function fncGetKTNM2(ByVal KTCD As String) As String
    Dim lKTCD As Long

    fncGetKTNM2 = ""
    Select Case KTCD
    Case "1": fncGetKTNM2 = "成型"
    Case "2": fncGetKTNM2 = "冷却"
    Case Else
        If IsNumeric(KTCD) Then
            lKTCD = CLng(Val(KTCD))
            If lKTCD >= 21 Then
                fncGetKTNM2 = "冷却"
            ElseIf lKTCD > 0 Then
                fncGetKTNM2 = "成型"
            End If
        End If
    End Select
End Function


