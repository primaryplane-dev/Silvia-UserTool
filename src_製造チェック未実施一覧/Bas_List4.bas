Attribute VB_Name = "Bas_List4"
Option Explicit

Public Sub subPrepareEditSheet()
    Call subBeforeEdit
    Call subEditList4
    Call subAfterEdit
End Sub

Private Function fncGetTargetKTCD() As Long
    Dim lKTCD As Long

    lKTCD = Val(P_KTCD)
    If lKTCD >= 21 Then
        fncGetTargetKTCD = 21
    Else
        fncGetTargetKTCD = 1
    End If
End Function

Private Sub subInitialize(ByVal ST As Worksheet, ByVal HN As Worksheet)
    HN.Cells.Copy ST.Cells
    ST.Range("13:19").Delete
    ST.Select
End Sub

Private Sub subEditList4()
    Dim lTargetKTCD As Long
    Dim ST          As Worksheet
    Dim HN          As Worksheet
    Dim CN          As ADODB.Connection
    Dim RS          As ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lCol        As Long
    Dim strKey      As String
    Dim lKTRow      As Long
    Dim lSGRow      As Long
    Dim lKRRow      As Long
    Dim lBKRow      As Long
    Dim vCVNO       As Variant
    Dim strDateYMD  As String
    Dim strTimeHMM  As String
    Dim strTimeHMS  As String
    Dim strDateTime As String

    lTargetKTCD = fncGetTargetKTCD()
    Set ST = stList4
    Set HN = stHina4
    
    '雛型シート→編集シート
    Call subInitialize(ST, HN)
    
    '見だし
    ST.Cells(1, 2) = "ビスケットコンベアチェック表(" & fncGetKTNM2(CStr(lTargetKTCD)) & ")"
    ST.Cells(5, 3) = Format(P_DATE, "yyyy年m月d日")
    ST.Cells(5, 5) = P_HINM
    ST.Cells(1, 4) = "製造終了"
    
    'ＤＢ接続
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    'SBFP01:ビスケットコンベアチェックマスタ
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "   BFCVNO "
    strSQL = strSQL & "  ,BFCVNM "
    strSQL = strSQL & "  FROM LIBSMF17.SBFP01 "
    strSQL = strSQL & " WHERE BFDELT = '' "
    strSQL = strSQL & "   AND BFKTCD = " & lTargetKTCD
    strSQL = strSQL & " ORDER BY BFCVNO "
    
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    lRow = 12
    Do While Not RS.EOF
        HN.Range(HN.Cells(12, 1), HN.Cells(11, 14)).Copy ST.Cells(lRow, 1)
        ST.Cells(lRow, 2) = RS("BFCVNO")
        ST.Cells(lRow, 3) = RS("BFCVNM")
        lRow = lRow + 1
        RS.MoveNext
    Loop
    RS.Close
    
    lKTRow = lRow           '固定チェック項目
    lSGRow = lRow + 2       '作業者
    lKRRow = lSGRow + 2     '管理者
    lBKRow = lKRRow + 2     '備考欄
    
    '固定チェック項目、作業者、管理者、備考欄を挿入する
    HN.Range(HN.Cells(13, 1), HN.Cells(19, 14)).Copy ST.Cells(lKTRow, 1)
    Call subHideTemplateRows(ST, lKTRow, lKTRow + 6)
    ST.Rows(lSGRow + 1).Hidden = True
    ST.Rows(lKRRow + 1).Hidden = True
    
    '罫線を引く
    ST.Range(ST.Cells(12, 2), ST.Cells(lBKRow, 13)).Borders.LineStyle = xlContinuous
    
    'SBGP01:ビスケットコンベアチェック
    strSQL = ""
    strSQL = strSQL & "SELECT BG.* "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS1.TSSYKJ,'')='' THEN VA1.VASYKJ ELSE COALESCE(TS1.TSSYKJ,'') END AS SYKJ "       '担当者名
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS2.TSSYKJ,'')='' THEN VA2.VASYKJ ELSE COALESCE(TS2.TSSYKJ,'') END AS KKSYKJ "     '確認者名
    strSQL = strSQL & "  FROM LIBSMF17.SBGP01 AS BG "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 "
    strSQL = strSQL & "    ON TS1.TSDELT = '' "
    strSQL = strSQL & "   AND TS1.TSSYCD = BG.BGSGCD"
    strSQL = strSQL & "   AND TS1.TSTTKB = '1' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA1 "
    strSQL = strSQL & "    ON VA1.VAKYUK = '' "
    strSQL = strSQL & "   AND VA1.VASYCD = BG.BGSGCD"
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 "
    strSQL = strSQL & "    ON TS2.TSDELT = '' "
    strSQL = strSQL & "   AND TS2.TSSYCD = BG.BGCHKK "
    strSQL = strSQL & "   AND TS2.TSTTKB = '4' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA2 "
    strSQL = strSQL & "    ON VA2.VAKYUK = '' "
    strSQL = strSQL & "   AND VA2.VASYCD = BG.BGCHKK "
    strSQL = strSQL & " WHERE BG.BGDELT = '' "
    strSQL = strSQL & "   AND BG.BGSDAT = " & Format(P_DATE, "yyyymmdd")
    strSQL = strSQL & "   AND BG.BGHINO = " & Val(P_HINO)
    strSQL = strSQL & "   AND BG.BGKJNO = '" & P_KJNO & "'"
    strSQL = strSQL & "   AND BG.BGKTCD = " & lTargetKTCD
    strSQL = strSQL & " ORDER BY BG.BGDATE, BG.BGTIME "
    
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lCol = 3: strKey = ""
    Do While Not RS.EOF
        strDateYMD = fncFormatYMD(RS("BGDATE"))
        strTimeHMS = fncFormatHMS(RS("BGTIME"))
        strTimeHMM = fncFormatHMM(RS("BGTIME"))
        strDateTime = Trim$(strDateYMD & " " & strTimeHMS)
        If Not strKey = Replace$(strDateYMD, "/", "") & Replace$(strTimeHMS, ":", "") Then
            lCol = lCol + 1
            ST.Cells(9, lCol) = strTimeHMM
            ST.Cells(10, lCol) = strDateTime
            ST.Cells(11, lCol) = strDateTime
            strKey = Replace$(strDateYMD, "/", "") & Replace$(strTimeHMS, ":", "")
        End If
        vCVNO = RS("BGCVNO")
        If IsNull(vCVNO) Then
            lRow = 0
        ElseIf Trim$(CStr(vCVNO)) = "" Then
            lRow = 0
        Else
            lRow = fncFindRow(ST, CLng(Val(CStr(vCVNO))), lKTRow)
        End If
        If lRow > 0 Then
            ST.Cells(lRow, lCol) = fncGetKigo(RS("BGCKRT"))
            ST.Cells(lSGRow, lCol) = RS("SYKJ")
            ST.Cells(lSGRow + 1, lCol) = Format(RS("BGSGCD"), "00000000")
            ST.Cells(lKRRow, lCol) = RS("KKSYKJ")
            ST.Cells(lKRRow + 1, lCol) = Format(RS("BGCHKK"), "00000000")
            ST.Cells(lBKRow, lCol) = RS("BGBIKO")
        End If
        RS.MoveNext
    Loop
    
    RS.Close
    
    'ライン長、品質管理、最終確認者取得
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT "
    strSQL = strSQL & "       BGCHKL "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS1.TSSYKJ,'')='' THEN VA1.VASYKJ ELSE COALESCE(TS1.TSSYKJ,'') END AS KLSYKJ "             'ライン長名
    strSQL = strSQL & "     , BGCHKH "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS2.TSSYKJ,'')='' THEN VA2.VASYKJ ELSE COALESCE(TS2.TSSYKJ,'') END AS KHSYKJ "             '品管名
    strSQL = strSQL & "     , BGCHKS "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS3.TSSYKJ,'')='' THEN VA3.VASYKJ ELSE COALESCE(TS3.TSSYKJ,'') END AS KSSYKJ "             '最終確認者名
    strSQL = strSQL & "  FROM LIBSMF17.SBGP01 "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 "
    strSQL = strSQL & "    ON TS1.TSDELT = '' "
    strSQL = strSQL & "   AND TS1.TSSYCD = BGCHKL"
    strSQL = strSQL & "   AND TS1.TSTTKB = '2'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA1 "
    strSQL = strSQL & "    ON VA1.VAKYUK = '' "
    strSQL = strSQL & "   AND VA1.VASYCD = BGCHKL"
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 "
    strSQL = strSQL & "    ON TS2.TSDELT = '' "
    strSQL = strSQL & "   AND TS2.TSSYCD = BGCHKH "
    strSQL = strSQL & "   AND TS2.TSTTKB = '2'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA2 "
    strSQL = strSQL & "    ON VA2.VAKYUK = '' "
    strSQL = strSQL & "   AND VA2.VASYCD = BGCHKH "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS3 "
    strSQL = strSQL & "    ON TS3.TSDELT = '' "
    strSQL = strSQL & "   AND TS3.TSSYCD = BGCHKS "
    strSQL = strSQL & "   AND TS3.TSTTKB = '3'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA3 "
    strSQL = strSQL & "    ON VA3.VAKYUK = '' "
    strSQL = strSQL & "   AND VA3.VASYCD = BGCHKS "
    strSQL = strSQL & " WHERE BGDELT='' "
    strSQL = strSQL & "   AND BGSDAT=" & Val(Format(P_DATE, "yyyymmdd"))
    strSQL = strSQL & "   AND BGHINO=" & Val(P_HINO)
    strSQL = strSQL & "   AND BGKJNO='" & P_KJNO & "'"
    strSQL = strSQL & "   AND BGKTCD=" & lTargetKTCD
    strSQL = strSQL & "   AND BGFNCD='1' "
    strSQL = strSQL & "   AND (BGCHKL<>'' OR BGCHKH<>'' OR BGCHKS<>'')"

    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    If RS.RecordCount > 0 Then
        ST.Cells(6, 11) = RS("BGCHKS")
        ST.Cells(7, 11) = RS("KSSYKJ")
        ST.Cells(6, 12) = RS("BGCHKH")
        ST.Cells(7, 12) = RS("KHSYKJ")
        ST.Cells(6, 13) = RS("BGCHKL")
        ST.Cells(7, 13) = RS("KLSYKJ")
    End If
    
    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing
        
    ST.Cells(9, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set HN = Nothing
    Set ST = Nothing
End Sub

Private Function fncFormatYMD(ByVal vDate As Variant) As String
    Dim s As String

    fncFormatYMD = ""
    If IsNull(vDate) Then Exit Function

    s = Trim$(CStr(vDate))
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    s = Right$("00000000" & CStr(CLng(s)), 8)
    fncFormatYMD = Left$(s, 4) & "/" & Mid$(s, 5, 2) & "/" & Right$(s, 2)
End Function

Private Function fncFormatHMS(ByVal vTime As Variant) As String
    Dim s As String
    Dim hh As Long
    Dim mm As Long
    Dim ss As Long

    fncFormatHMS = ""
    If IsNull(vTime) Then Exit Function

    s = Trim$(CStr(vTime))
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    s = Right$("000000" & CStr(CLng(s)), 6)
    hh = CLng(Left$(s, 2))
    mm = CLng(Mid$(s, 3, 2))
    ss = CLng(Right$(s, 2))
    If hh > 23 Or mm > 59 Or ss > 59 Then Exit Function

    fncFormatHMS = Right$("0" & CStr(hh), 2) & ":" & Right$("0" & CStr(mm), 2) & ":" & Right$("0" & CStr(ss), 2)
End Function

Private Function fncFormatHMM(ByVal vTime As Variant) As String
    Dim s As String
    Dim hh As Long
    Dim mm As Long

    fncFormatHMM = ""
    If IsNull(vTime) Then Exit Function

    s = Trim$(CStr(vTime))
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    s = Right$("000000" & CStr(CLng(s)), 6)
    hh = CLng(Left$(s, 2))
    mm = CLng(Mid$(s, 3, 2))
    If hh > 23 Or mm > 59 Then Exit Function

    fncFormatHMM = CStr(hh) & ":" & Right$("0" & CStr(mm), 2)
End Function

Private Function fncFindRow(ByVal ST As Worksheet, ByVal CVNO As Long, ByVal lKTRow As Long) As Long
    Dim lRow        As Long
    Dim vNo         As Variant
    Dim sNo         As String

    fncFindRow = 0
    If CVNO >= 901 Then
        fncFindRow = lKTRow + CVNO - 901
    Else
        lRow = 12
        Do While lRow < lKTRow
            vNo = ST.Cells(lRow, 2).Value
            sNo = Trim$(CStr(vNo))
            If sNo <> "" Then
                sNo = StrConv(sNo, vbNarrow)
                If IsNumeric(sNo) Then
                    If CLng(Val(sNo)) = CVNO Then
                        fncFindRow = lRow
                        Exit Do
                    End If
                End If
            End If
            lRow = lRow + 1
        Loop

        ' 固定チェック項目が 1～7 で保持されるデータに対応
        If fncFindRow = 0 Then
            If CVNO >= 1 And CVNO <= 7 Then
                fncFindRow = lKTRow + CVNO - 1
            End If
        End If
    End If

End Function

Private Function fncGetKigo(ByVal i_KigoCD As String) As String
    ' 表示は stList3 と同じ 3 記号(×/○/－)に統一し、旧データのコード値(3/4/5)も互換吸収する
    Select Case i_KigoCD
    Case "0", "4": fncGetKigo = "×"
    Case "1", "5": fncGetKigo = "○"
    Case "2", "3": fncGetKigo = "－"
    Case Else: fncGetKigo = i_KigoCD
    End Select
End Function

Private Sub subHideTemplateRows(ByVal ST As Worksheet, ByVal lStartRow As Long, ByVal lEndRow As Long)
    Dim lRow        As Long
    Dim sName       As String

    For lRow = lStartRow To lEndRow
        sName = fncNormalizeName(CStr(ST.Cells(lRow, 3).Value))
        If sName = "" Then sName = fncNormalizeName(CStr(ST.Cells(lRow, 2).Value))
        If sName = "" Then sName = fncNormalizeName(CStr(ST.Cells(lRow, 5).Value))
        If sName = "異音確認" Or sName = "動作確認(ノッキング)" Then
            ST.Rows(lRow).Hidden = True
        End If
    Next lRow
End Sub

Private Function fncNormalizeName(ByVal sValue As String) As String
    Dim s As String

    s = Trim$(sValue)
    s = Replace$(s, vbTab, "")
    s = Replace$(s, "　", "")
    s = Replace$(s, " ", "")
    s = Replace$(s, "（", "(")
    s = Replace$(s, "）", ")")
    fncNormalizeName = s
End Function

