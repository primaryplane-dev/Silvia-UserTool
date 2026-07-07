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
    Set ST = stList4

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
        strSQL = strSQL & "     ,BGCHKL = " & fncGetApproverCode(ST, 13)
        strSQL = strSQL & "     ,BGCHKH = " & fncGetApproverCode(ST, 12)
        strSQL = strSQL & "     ,BGCHKS = " & fncGetApproverCode(ST, 11)
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

Private Function fncGetApproverCode(ByVal ST As Worksheet, ByVal lCol As Long) As Long
    Dim vCode As Variant
    Dim sName As String

    fncGetApproverCode = 0

    'ÉRÅ[ÉhçsÅi6çsñ⁄ÅjÇóDêÊ
    vCode = ST.Cells(6, lCol).Value
    If IsNumeric(vCode) Then
        If Val(vCode) > 0 Then
            fncGetApproverCode = Val(vCode)
            Exit Function
        End If
    End If

    'å›ä∑: ÉRÅ[ÉhÇ™7çsñ⁄Ç…ì¸Ç¡ÇƒÇ¢ÇÈèÍçá
    vCode = ST.Cells(7, lCol).Value
    If IsNumeric(vCode) Then
        If Val(vCode) > 0 Then
            fncGetApproverCode = Val(vCode)
            Exit Function
        End If
    End If

    'éÅñºÇÃÇ›ì¸óÕÇ≥ÇÍÇƒÇ¢ÇÈèÍçáÇÕé–àıÉRÅ[ÉhÇãtà¯Ç´Ç∑ÇÈ
    sName = Trim$(CStr(ST.Cells(7, lCol).Value))
    If sName <> "" Then
        fncGetApproverCode = fncGetSYCDByName(sName)
    End If
End Function

Private Function fncGetSYCDByName(ByVal iSYNM As String) As Long
    Dim CN      As ADODB.Connection
    Dim RS      As ADODB.Recordset
    Dim strSQL  As String

    fncGetSYCDByName = 0
    If Trim$(iSYNM) = "" Then Exit Function

    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset

    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString

    strSQL = ""
    strSQL = strSQL & "SELECT TSSYCD "
    strSQL = strSQL & "  FROM LIBSMF17.STSP01 "
    strSQL = strSQL & " WHERE TSDELT='' "
    strSQL = strSQL & "   AND TSSYKJ='" & Replace(iSYNM, "'", "''") & "' "
    strSQL = strSQL & " FETCH FIRST 1 ROW ONLY"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If Not RS.EOF Then
        fncGetSYCDByName = Val(RS("TSSYCD"))
    End If
    RS.Close

    If fncGetSYCDByName = 0 Then
        strSQL = ""
        strSQL = strSQL & "SELECT VASYCD "
        strSQL = strSQL & "  FROM LIBBMF.BVAP01 "
        strSQL = strSQL & " WHERE VAKYUK='' "
        strSQL = strSQL & "   AND VASYKJ='" & Replace(iSYNM, "'", "''") & "' "
        strSQL = strSQL & " FETCH FIRST 1 ROW ONLY"

        RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
        If Not RS.EOF Then
            fncGetSYCDByName = Val(RS("VASYCD"))
        End If
        RS.Close
    End If

    Set RS = Nothing
    CN.Close
    Set CN = Nothing
End Function
