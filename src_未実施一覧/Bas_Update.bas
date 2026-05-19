Option Explicit

Public Sub subUpdate2()
    Dim strSQL      As String
    Dim CN          As ADODB.Connection
    Dim lResult     As Long

    'ÇcÇaê⁄ë±
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    'è≥îFÇÃÇ›
    If stList2.Cells(6, 15) = "1" Then
        strSQL = ""
        strSQL = strSQL & "UPDATE LIBSMF17.SBAP01 SET"
        strSQL = strSQL & "      BAUUSR='" & P_SYCD & "'"
        strSQL = strSQL & "     ,BAUPGM='" & P_PGM & "'"
        strSQL = strSQL & "     ,BAUNTH = TO_CHAR(current timestamp, 'YYYYMMDD')"
        strSQL = strSQL & "     ,BAUTIM = TO_CHAR(current timestamp, 'HH24MISS')"
        strSQL = strSQL & "     ,BACHKL = " & Val(stList2.Cells(6, 14))
        strSQL = strSQL & "     ,BACHKH = " & Val(stList2.Cells(6, 13))
        strSQL = strSQL & "     ,BACHKS = " & Val(stList2.Cells(6, 12))
        strSQL = strSQL & " WHERE BADELT='' "
        strSQL = strSQL & "    AND BADATE=" & Val(Format(P_DATE, "yyyymmdd"))
        strSQL = strSQL & "    AND BAHINO='" & P_HINO & "'"
        strSQL = strSQL & "    AND BAKJNO='" & P_KJNO & "'"
        CN.Execute strSQL, lResult, &H80
    End If
    'ÇcÇaêÿíf
    CN.Close: Set CN = Nothing
End Sub

Public Sub subUpdate3()
    Dim strSQL      As String
    Dim CN          As ADODB.Connection
    Dim lResult     As Long

    'ÇcÇaê⁄ë±
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    'è≥îFÇÃÇ›
    If stList3.Cells(1, 17) = "1" Then
        strSQL = ""
        strSQL = strSQL & "UPDATE LIBSMF17.SBEP01 SET"
        strSQL = strSQL & "      BEUUSR='" & P_SYCD & "'"
        strSQL = strSQL & "     ,BEUPGM='" & P_PGM & "'"
        strSQL = strSQL & "     ,BEUNTH = TO_CHAR(current timestamp, 'YYYYMMDD')"
        strSQL = strSQL & "     ,BEUTIM = TO_CHAR(current timestamp, 'HH24MISS')"
        strSQL = strSQL & "     ,BECHKL = " & Val(stList3.Cells(5, 16))
        strSQL = strSQL & "     ,BECHKH = " & Val(stList3.Cells(5, 15))
        strSQL = strSQL & "     ,BECHKS = " & Val(stList3.Cells(5, 14))
        strSQL = strSQL & " WHERE BEDELT='' "
        strSQL = strSQL & "    AND BESDAT=" & Val(Format(P_DATE, "yyyymmdd"))
        strSQL = strSQL & "    AND BEHINO='" & P_HINO & "'"
        strSQL = strSQL & "    AND BEKJNO='" & P_KJNO & "'"
        strSQL = strSQL & "    AND BESHBU=" & P_SHBU
        If P_KTCD = "1" Then
            strSQL = strSQL & "    AND BEKTCD >= 1 AND BEKTCD <= 20 "
        ElseIf P_KTCD = "2" Then
            strSQL = strSQL & "    AND BEKTCD >= 21 "
        End If
        CN.Execute strSQL, lResult, &H80
    End If
    'ÇcÇaêÿíf
    CN.Close: Set CN = Nothing
End Sub

Public Sub subUpdate4()
    Dim strSQL      As String
    Dim CN          As ADODB.Connection
    Dim lResult     As Long
    Dim lTargetKTCD As Long
    Dim ST          As Worksheet

    lTargetKTCD = IIf(Val(P_KTCD) >= 21, 21, 1)
    On Error Resume Next
    If lTargetKTCD = 21 Then
        Set ST = ThisWorkbook.Worksheets("stList5")
    Else
        Set ST = stList4
    End If
    On Error GoTo 0
    If ST Is Nothing Then Set ST = stList4

    'ÇcÇaê⁄ë±
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    'è≥îFÇÃÇ›
    If ST.Cells(1, 14) = "1" Then
        strSQL = ""
        strSQL = strSQL & "UPDATE LIBSMF17.SBGP01 SET"
        strSQL = strSQL & "      BGUUSR='" & P_SYCD & "'"
        strSQL = strSQL & "     ,BGUPGM='" & P_PGM & "'"
        strSQL = strSQL & "     ,BGUNTH = TO_CHAR(current timestamp, 'YYYYMMDD')"
        strSQL = strSQL & "     ,BGUTIM = TO_CHAR(current timestamp, 'HH24MISS')"
        strSQL = strSQL & "     ,BGCHKL = " & Val(ST.Cells(6, 13))
        strSQL = strSQL & "     ,BGCHKH = " & Val(ST.Cells(6, 12))
        strSQL = strSQL & "     ,BGCHKS = " & Val(ST.Cells(6, 11))
        strSQL = strSQL & " WHERE BGDELT='' "
        strSQL = strSQL & "    AND BGSDAT=" & Val(Format(P_DATE, "yyyymmdd"))
        strSQL = strSQL & "    AND BGHINO='" & P_HINO & "'"
        strSQL = strSQL & "    AND BGKJNO='" & P_KJNO & "'"
        strSQL = strSQL & "    AND BGKTCD=" & lTargetKTCD
        CN.Execute strSQL, lResult, &H80
    End If
    'ÇcÇaêÿíf
    CN.Close: Set CN = Nothing
    Set ST = Nothing
End Sub
