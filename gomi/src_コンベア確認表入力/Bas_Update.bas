Attribute VB_Name = "Bas_Update"
Option Explicit

Public Sub subUpdate()
    Dim strSQL      As String
    Dim CN          As New ADODB.Connection
    Dim lResult     As Long
    Dim lRow        As Long
    Dim ST          As Worksheet: Set ST = stList
    
    'ÇcÇaê⁄ë±
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    ' 1. çÌèúèàóù
    ' stWorkÇÃ14óÒñ⁄Ç™ "X" ÇÃÇ‡ÇÃÇçÌèú
    lRow = 1
    Do While Not stWork.Cells(lRow, 1) = ""
        If Not stWork.Cells(lRow, 1) = stWork.Cells(lRow, 2) Or stWork.Cells(lRow, 14) = "X" Then
            strSQL = ""
            strSQL = strSQL & " UPDATE LIBSMF17.SBGP01 SET "
            strSQL = strSQL & "      BGUUSR='" & P_SYCD & "'"
            strSQL = strSQL & "     ,BGUPGM='" & P_PGM & "'"
            strSQL = strSQL & "     ,BGUNTH = TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "     ,BGUTIM = TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & "     ,BGDELT = 'X'"
            strSQL = strSQL & "  WHERE BGDELT='' "
            strSQL = strSQL & "    AND BGSDAT=" & Val(Format(P_DATE, "yyyymmdd"))
            strSQL = strSQL & "    AND BGHINO='" & P_HINO & "'"
            strSQL = strSQL & "    AND BGKJNO='" & P_KJNO & "'"
            strSQL = strSQL & "    AND BGDATE=" & Val(Left(stWork.Cells(lRow, 2), 8))
            strSQL = strSQL & "    AND BGTIME=" & Val(Right(stWork.Cells(lRow, 2), 6))
            strSQL = strSQL & "    AND BGKTCD='" & stWork.Cells(lRow, 3) & "'"
            strSQL = strSQL & "    AND BGCVNO=" & Val(stWork.Cells(lRow, 4))
            CN.Execute strSQL, lResult, &H80
        End If
        lRow = lRow + 1
    Loop
    
    ' 2. çXêVÅEí«â¡èàóù
    lRow = 1
    Do While Not stWork.Cells(lRow, 1) = ""
        If Not stWork.Cells(lRow, 14) = "X" Then
            ' UPDATEééçs
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBSMF17.SBGP01 SET"
            strSQL = strSQL & "      BGUUSR='" & P_SYCD & "'"
            strSQL = strSQL & "     ,BGUPGM='" & P_PGM & "'"
            strSQL = strSQL & "     ,BGUNTH = TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "     ,BGUTIM = TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & "     ,BGCKRT = '" & stWork.Cells(lRow, 6) & "'"
            strSQL = strSQL & "     ,BGBIKO = '" & stWork.Cells(lRow, 11) & "' "
            strSQL = strSQL & "     ,BGTMCD = '" & stWork.Cells(lRow, 12) & "' "
            strSQL = strSQL & "     ,BGFNCD = '" & stWork.Cells(lRow, 13) & "' "
            strSQL = strSQL & "     ,BGCHKK = " & Val(stWork.Cells(lRow, 9))
            strSQL = strSQL & "     ,BGCHKL = " & Val(ST.Cells(5, 16))
            strSQL = strSQL & "     ,BGCHKH = " & Val(ST.Cells(5, 15))
            strSQL = strSQL & "     ,BGCHKS = " & Val(ST.Cells(5, 14))
            strSQL = strSQL & "  WHERE BGDELT='' "
            strSQL = strSQL & "    AND BGSDAT=" & Val(Format(P_DATE, "yyyymmdd"))
            strSQL = strSQL & "    AND BGHINO='" & P_HINO & "'"
            strSQL = strSQL & "    AND BGKJNO='" & P_KJNO & "'"
            strSQL = strSQL & "    AND BGDATE=" & Val(Left(stWork.Cells(lRow, 1), 8))
            strSQL = strSQL & "    AND BGTIME=" & Val(Right(stWork.Cells(lRow, 1), 6))
            strSQL = strSQL & "    AND BGKTCD='" & stWork.Cells(lRow, 3) & "'"
            strSQL = strSQL & "    AND BGCVNO=" & Val(stWork.Cells(lRow, 4))
            
            CN.Execute strSQL, lResult, &H80
            
            ' INSERT (êVãK)
            If lResult = 0 Then
                strSQL = ""
                strSQL = strSQL & " INSERT INTO LIBSMF17.SBGP01"
                strSQL = strSQL & "         ("
                strSQL = strSQL & "              BGDELT"
                strSQL = strSQL & "             ,BGCUSR"
                strSQL = strSQL & "             ,BGCPGM"
                strSQL = strSQL & "             ,BGCNTH"
                strSQL = strSQL & "             ,BGCTIM"
                strSQL = strSQL & "             ,BGSDAT"
                strSQL = strSQL & "             ,BGHINO"
                strSQL = strSQL & "             ,BGHINM"
                strSQL = strSQL & "             ,BGKJNO"
                strSQL = strSQL & "             ,BGKJNM"
                strSQL = strSQL & "             ,BGDATE"
                strSQL = strSQL & "             ,BGTIME"
                strSQL = strSQL & "             ,BGTMCD"
                strSQL = strSQL & "             ,BGKTCD"
                strSQL = strSQL & "             ,BGCVNO"
                strSQL = strSQL & "             ,BGCKRT"
                strSQL = strSQL & "             ,BGFNCD"
                strSQL = strSQL & "             ,BGBIKO"
                strSQL = strSQL & "             ,BGSGCD"
                strSQL = strSQL & "             ,BGCHKK"
                strSQL = strSQL & "             ,BGCHKL"
                strSQL = strSQL & "             ,BGCHKH"
                strSQL = strSQL & "             ,BGCHKS"
                strSQL = strSQL & "         ) VALUES ("
                strSQL = strSQL & "              ''"
                strSQL = strSQL & "             ,'" & P_SYCD & "'"
                strSQL = strSQL & "             ,'" & P_PGM & "'"
                strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'YYYYMMDD')"
                strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'HH24MISS')"
                strSQL = strSQL & "             ," & Val(Format(P_DATE, "yyyymmdd"))
                strSQL = strSQL & "             ,'" & P_HINO & "'"
                strSQL = strSQL & "             ,'" & P_HINM & "'"
                strSQL = strSQL & "             ,'" & P_KJNO & "'"
                strSQL = strSQL & "             ,'" & P_KJNM & "'"
                strSQL = strSQL & "             ," & Val(Left(stWork.Cells(lRow, 1), 8))
                strSQL = strSQL & "             ," & Val(Right(stWork.Cells(lRow, 1), 6))
                strSQL = strSQL & "             ,'" & stWork.Cells(lRow, 12) & "'"
                strSQL = strSQL & "             ,'" & stWork.Cells(lRow, 3) & "'"
                strSQL = strSQL & "             ," & Val(stWork.Cells(lRow, 4))
                strSQL = strSQL & "             ,'" & stWork.Cells(lRow, 6) & "'"
                strSQL = strSQL & "             ,'" & stWork.Cells(lRow, 13) & "'"
                strSQL = strSQL & "             ,'" & stWork.Cells(lRow, 11) & "'"
                strSQL = strSQL & "             ," & Val(stWork.Cells(lRow, 7))
                strSQL = strSQL & "             ," & Val(stWork.Cells(lRow, 9))
                strSQL = strSQL & "             ," & Val(ST.Cells(5, 16))
                strSQL = strSQL & "             ," & Val(ST.Cells(5, 15))
                strSQL = strSQL & "             ," & Val(ST.Cells(5, 14))
                strSQL = strSQL & "         )"
                
                CN.Execute strSQL, lResult, &H80
            End If
        End If
        lRow = lRow + 1
    Loop
    CN.Close: Set CN = Nothing
    Set ST = Nothing
End Sub
