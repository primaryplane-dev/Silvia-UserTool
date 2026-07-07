Option Explicit

Public Sub subProcessEditList()
    Call subBeforeEdit
    Call subEditList3
    Call subAfterEdit
End Sub

Private Sub subPerformEditList()
    Dim ST          As Worksheet: Set ST = stList3
    Dim CN          As ADODB.Connection
    Dim RS          As ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim startRow    As Long
    Dim lCol        As Long
    Dim lCol2       As Long
    Dim bEnd        As Boolean
    Dim strKey      As String
    Dim Cnt         As Integer
    Dim lEndRow     As Long
    Dim lSGRow      As Long
    Dim lKRRow      As Long
    Dim lBKRow      As Long

    '雛型シート→編集シート
    Call subInitialize

    '見だし
    ST.Cells(1, 2) = fncGetSHNM(P_SHBU) & "備品チェック表(" & fncGetKTNM2(P_KTCD) & ")"
    ST.Cells(4, 5) = Format(P_DATE, "yyyy年m月d日")
    ST.Cells(4, 8) = P_HINM
    ST.Cells(1, 7) = "製造終了"
End Sub

Private Sub subInitialize()
    stHina3.Cells.Copy stList3.Cells
    stList3.Range("12:16").Delete
    stList3.Select
End Sub

Private Sub subEditList3()
    Dim ST          As Worksheet: Set ST = stList3
    Dim CN          As ADODB.Connection
    Dim RS          As ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim startRow    As Long
    Dim lCol        As Long
    Dim lCol2       As Long
    Dim bEnd        As Boolean
    Dim strKey      As String
    Dim Cnt         As Integer
    Dim lEndRow     As Long
    Dim lSGRow      As Long
    Dim lKRRow      As Long
    Dim lBKRow      As Long
    
    '雛型シート→編集シート
    Call subInitialize
    
    '見だし
    ST.Cells(1, 2) = fncGetSHNM(P_SHBU) & "備品チェック表(" & fncGetKTNM2(P_KTCD) & ")"
    ST.Cells(4, 5) = Format(P_DATE, "yyyy年m月d日")
    ST.Cells(4, 8) = P_HINM
    ST.Cells(1, 7) = "製造終了"
    
    'ＤＢ接続
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    'SBDP01:ビスケット備品チェック項目マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "   BDKTCD "
    strSQL = strSQL & "  ,BDCKNO "
    strSQL = strSQL & "  ,BDCKCT "
    strSQL = strSQL & "  ,COALESCE(BDSGCD,'') AS BDSGCD "
    strSQL = strSQL & "  FROM LIBSMF17.SBDP01 "
    strSQL = strSQL & " WHERE BDDELT = '' "
    strSQL = strSQL & "   AND BDSHBU =" & P_SHBU
    If P_KTCD = "1" Then
        strSQL = strSQL & "   AND BDKTCD >=1 AND BDKTCD <= 20 "
    ElseIf P_KTCD = "2" Then
        strSQL = strSQL & "   AND BDKTCD >=21 "
    End If
    strSQL = strSQL & " ORDER BY BDKTCD, BDCKNO "
    
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    lRow = 11
    Do While Not RS.EOF
        stHina3.Range(stHina3.Cells(11, 1), stHina3.Cells(11, 17)).Copy ST.Cells(lRow, 1)
        ST.Cells(lRow, 2) = fncGetKTNM(RS("BDKTCD"))
        ST.Cells(lRow, 3) = RS("BDKTCD")
        ST.Cells(lRow, 4) = CStr(RS("BDCKNO"))
        ST.Cells(lRow, 5) = RS("BDCKCT")
        ST.Cells(lRow, 6) = RS("BDSGCD")
        lRow = lRow + 1
        RS.MoveNext
    Loop
    RS.Close
    
    lSGRow = lRow
    lKRRow = lRow + 2
    lBKRow = lRow + 4
    
    '作業者、管理者、備考欄を挿入する
    stHina3.Range(stHina3.Cells(12, 1), stHina3.Cells(16, 17)).Copy ST.Cells(lSGRow, 1)
    ST.Rows(lSGRow + 1).Hidden = True
    ST.Rows(lKRRow + 1).Hidden = True
    
    Application.DisplayAlerts = False
    '同一の作業工程をマージする
    strKey = "": startRow = 11
    For lRow = 11 To lSGRow - 1
        If Not strKey = ST.Cells(lRow, 3) Then
            If Not strKey = "" Then
                ST.Range(ST.Cells(startRow, 2), ST.Cells(lRow - 1, 2)).Merge
            End If
            startRow = lRow
            strKey = ST.Cells(lRow, 3)
        End If
    Next
    If Not strKey = "" Then
        ST.Range(ST.Cells(startRow, 2), ST.Cells(lSGRow - 1, 2)).Merge
    End If
    Application.DisplayAlerts = True
    
    '罫線を引く
    ST.Range(ST.Cells(11, 2), ST.Cells(lBKRow, 16)).Borders.LineStyle = xlContinuous
    
    'SBEP01:ビスケット備品チェック
    strSQL = ""
    strSQL = strSQL & "SELECT BE.* "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS1.TSSYKJ,'')='' THEN VA1.VASYKJ ELSE COALESCE(TS1.TSSYKJ,'') END AS SYKJ "       '担当者名
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS2.TSSYKJ,'')='' THEN VA2.VASYKJ ELSE COALESCE(TS2.TSSYKJ,'') END AS KKSYKJ "     '確認者名
    strSQL = strSQL & "  FROM LIBSMF17.SBEP01 AS BE "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 "
    strSQL = strSQL & "    ON TS1.TSDELT = '' "
    strSQL = strSQL & "   AND TS1.TSSYCD = BE.BESGCD"
    strSQL = strSQL & "   AND TS1.TSTTKB = '1' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA1 "
    strSQL = strSQL & "    ON VA1.VAKYUK = '' "
    strSQL = strSQL & "   AND VA1.VASYCD = BE.BESGCD"
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 "
    strSQL = strSQL & "    ON TS2.TSDELT = '' "
    strSQL = strSQL & "   AND TS2.TSSYCD = BE.BECHKK "
    strSQL = strSQL & "   AND TS2.TSTTKB = '4' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA2 "
    strSQL = strSQL & "    ON VA2.VAKYUK = '' "
    strSQL = strSQL & "   AND VA2.VASYCD = BE.BECHKK "
    strSQL = strSQL & " WHERE BE.BEDELT = '' "
    strSQL = strSQL & "   AND BE.BESDAT = " & Format(P_DATE, "yyyymmdd")
    strSQL = strSQL & "   AND BE.BEHINO = " & Val(P_HINO)
    strSQL = strSQL & "   AND BE.BEKJNO = '" & P_KJNO & "'"
    strSQL = strSQL & "   AND BE.BESHBU = " & P_SHBU
    If P_KTCD = "1" Then
        strSQL = strSQL & "   AND BE.BEKTCD >= 1 AND BE.BEKTCD <= 20 "
    ElseIf P_KTCD = "2" Then
        strSQL = strSQL & "   AND BE.BEKTCD >= 21 "
    End If
    strSQL = strSQL & " ORDER BY BE.BEDATE, BE.BETIME "
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lCol = 6: strKey = ""
    Do While Not RS.EOF
        If Not strKey = RS("BEDATE") & Format(RS("BETIME"), "000000") Then
            lCol = lCol + 1
            ST.Cells(8, lCol) = Format(CDate(Format(RS("BETIME"), "00:00:00")), "H:MM")
            ST.Cells(9, lCol) = Format(RS("BEDATE"), "0000/00/00") & " " & Format(RS("BETIME"), "00:00:00")
            ST.Cells(10, lCol) = Format(RS("BEDATE"), "0000/00/00") & " " & Format(RS("BETIME"), "00:00:00")
            strKey = RS("BEDATE") & Format(RS("BETIME"), "000000")
        End If
        lRow = fncFindRow(RS("BEKTCD"), CLng(RS("BECKNO")))
        If lRow > 0 Then
            If RS("BECKRT") = "0" Then ST.Cells(lRow, lCol) = "×"
            If RS("BECKRT") = "1" Then ST.Cells(lRow, lCol) = "○"
            If RS("BECKRT") = "2" Then ST.Cells(lRow, lCol) = "－"
            ST.Cells(lSGRow, lCol) = RS("SYKJ")
            ST.Cells(lSGRow + 1, lCol) = Format(RS("BESGCD"), "00000000")
            ST.Cells(lKRRow, lCol) = RS("KKSYKJ")
            ST.Cells(lKRRow + 1, lCol) = Format(RS("BECHKK"), "00000000")
            ST.Cells(lBKRow, 7) = RS("BEBIKO")
        End If
        RS.MoveNext
    Loop
    
    RS.Close
    
    'ライン長、品質管理、最終確認者取得
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT "
    strSQL = strSQL & "       BECHKL "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS1.TSSYKJ,'')='' THEN VA1.VASYKJ ELSE COALESCE(TS1.TSSYKJ,'') END AS KLSYKJ "             'ライン長名
    strSQL = strSQL & "     , BECHKH "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS2.TSSYKJ,'')='' THEN VA2.VASYKJ ELSE COALESCE(TS2.TSSYKJ,'') END AS KHSYKJ "             '品管名
    strSQL = strSQL & "     , BECHKS "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS3.TSSYKJ,'')='' THEN VA3.VASYKJ ELSE COALESCE(TS3.TSSYKJ,'') END AS KSSYKJ "             '最終確認者名
    strSQL = strSQL & "  FROM LIBSMF17.SBEP01 "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 "
    strSQL = strSQL & "    ON TS1.TSDELT = '' "
    strSQL = strSQL & "   AND TS1.TSSYCD = BECHKL"
    strSQL = strSQL & "   AND TS1.TSTTKB = '2'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA1 "
    strSQL = strSQL & "    ON VA1.VAKYUK = '' "
    strSQL = strSQL & "   AND VA1.VASYCD = BECHKL"
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 "
    strSQL = strSQL & "    ON TS2.TSDELT = '' "
    strSQL = strSQL & "   AND TS2.TSSYCD = BECHKH "
    strSQL = strSQL & "   AND TS2.TSTTKB = '2'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA2 "
    strSQL = strSQL & "    ON VA2.VAKYUK = '' "
    strSQL = strSQL & "   AND VA2.VASYCD = BECHKH "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS3 "
    strSQL = strSQL & "    ON TS3.TSDELT = '' "
    strSQL = strSQL & "   AND TS3.TSSYCD = BECHKS "
    strSQL = strSQL & "   AND TS3.TSTTKB = '3'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA3 "
    strSQL = strSQL & "    ON VA3.VAKYUK = '' "
    strSQL = strSQL & "   AND VA3.VASYCD = BECHKS "
    strSQL = strSQL & " WHERE BEDELT='' "
    strSQL = strSQL & "   AND BESDAT=" & Val(Format(P_DATE, "yyyymmdd"))
    strSQL = strSQL & "   AND BEHINO=" & Val(P_HINO)
    strSQL = strSQL & "   AND BEKJNO='" & P_KJNO & "'"
    strSQL = strSQL & "   AND BESHBU = " & P_SHBU
    If P_KTCD = "1" Then
        strSQL = strSQL & "   AND BEKTCD >= 1 AND BEKTCD <= 20 "
    ElseIf P_KTCD = "2" Then
        strSQL = strSQL & "   AND BEKTCD >= 21 "
    End If
    strSQL = strSQL & "   AND BEFNCD='1' "
    strSQL = strSQL & "   AND (BECHKL<>'' OR BECHKH<>'' OR BECHKS<>'')"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    If RS.RecordCount > 0 Then
        ST.Cells(5, 14) = RS("BECHKS")
        ST.Cells(6, 14) = RS("KSSYKJ")
        ST.Cells(5, 15) = RS("BECHKH")
        ST.Cells(6, 15) = RS("KHSYKJ")
        ST.Cells(5, 16) = RS("BECHKL")
        ST.Cells(6, 16) = RS("KLSYKJ")
    End If
    
    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing
        
    ST.Cells(8, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set ST = Nothing
End Sub

Private Function fncGetKTNM(ByVal KTCD As String) As String
    fncGetKTNM = ""
    Select Case KTCD
    Case "1": fncGetKTNM = "成型"
    Case "2": fncGetKTNM = "スクリーン"
    Case "21": fncGetKTNM = "冷却"
    End Select
End Function

Private Function fncFindRow(ByVal KTCD As String, ByVal CKNO As Long) As Long
    Dim lRow        As Long
    Dim ST          As Worksheet: Set ST = stList3

    fncFindRow = 0
    
    lRow = 11
    Do While Not ST.Cells(lRow, 3) = ""
        If ST.Cells(lRow, 3) = KTCD And CLng(ST.Cells(lRow, 4)) = CKNO Then
            fncFindRow = lRow
            Exit Do
        End If
        lRow = lRow + 1
    Loop
    
    Set ST = Nothing
End Function

Public Sub subShowList()
    ' 必要な処理をここに記述
    MsgBox "リストを表示します"
End Sub

