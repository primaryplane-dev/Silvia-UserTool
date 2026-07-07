Attribute VB_Name = "Bas_List2"
Option Explicit

Public Sub subMain2()
    Call subBeforeEdit
    Call subEditList2
    Call subAfterEdit
End Sub

Private Sub subInitialize()
    stHina2.Cells.Copy stList2.Cells
    stList2.Select
End Sub

Private Sub subEditList2()
    Dim ST          As Worksheet: Set ST = stList2
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
    
    '雛型シート→編集シート
    Call subInitialize
    
    '見だし
    ST.Cells(4, 3) = Format(P_DATE, "yyyy年m月d日")
    ST.Cells(5, 3) = P_HINM
    ST.Cells(1, 8) = "生産終了"
        
    'ＤＢ接続
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    bEnd = False
    'SBAP01:ビスケット分割重量チェック
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT CASE WHEN BA1.BATNTM = 0 THEN 99999999999999 ELSE BA1.BATNTM END AS TNTM "
    strSQL = strSQL & "     , CASE WHEN BA1.BAKRTM = 0 THEN 99999999999999 ELSE BA1.BAKRTM END AS KRTM "
    strSQL = strSQL & "     , BA1.* "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS1.TSSYKJ,'')='' THEN VA1.VASYKJ ELSE COALESCE(TS1.TSSYKJ,'') END AS SYKJ "    '担当者名
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS2.TSSYKJ,'')='' THEN VA2.VASYKJ ELSE COALESCE(TS2.TSSYKJ,'') END AS KLSYKJ "  'ライン長名
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS3.TSSYKJ,'')='' THEN VA3.VASYKJ ELSE COALESCE(TS3.TSSYKJ,'') END AS KHSYKJ "  '品管名
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS4.TSSYKJ,'')='' THEN VA4.VASYKJ ELSE COALESCE(TS4.TSSYKJ,'') END AS KSSYKJ "  '最終確認者名
    strSQL = strSQL & "  FROM LIBSMF17.SBAP01 AS BA1 "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 "
    strSQL = strSQL & "    ON TS1.TSDELT = '' "
    strSQL = strSQL & "   AND TS1.TSSYCD = BASGCD"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA1 "
    strSQL = strSQL & "    ON VA1.VAKYUK = '' "
    strSQL = strSQL & "   AND VA1.VASYCD = BASGCD"
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 "
    strSQL = strSQL & "    ON TS2.TSDELT = '' "
    strSQL = strSQL & "   AND TS2.TSSYCD = BACHKL "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA2 "
    strSQL = strSQL & "    ON VA2.VAKYUK = '' "
    strSQL = strSQL & "   AND VA2.VASYCD = BACHKL "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS3 "
    strSQL = strSQL & "    ON TS3.TSDELT = '' "
    strSQL = strSQL & "   AND TS3.TSSYCD = BACHKH "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA3 "
    strSQL = strSQL & "    ON VA3.VAKYUK = '' "
    strSQL = strSQL & "   AND VA3.VASYCD = BACHKH "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS4 "
    strSQL = strSQL & "    ON TS4.TSDELT = '' "
    strSQL = strSQL & "   AND TS4.TSSYCD = BACHKS "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA4 "
    strSQL = strSQL & "    ON VA4.VAKYUK = '' "
    strSQL = strSQL & "   AND VA4.VASYCD = BACHKS "
    strSQL = strSQL & " WHERE BA1.BADELT='' "
    strSQL = strSQL & "   AND BA1.BADATE=" & Val(Format(P_DATE, "yyyymmdd"))
    strSQL = strSQL & "   AND BA1.BAHINO=" & Val(P_HINO)
    strSQL = strSQL & "   AND BA1.BAKJNO='" & P_KJNO & "'"
    strSQL = strSQL & "ORDER BY TNTM, KRTM "
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lRow = 10
    strKey = "": Cnt = 0: startRow = 10
    Do While Not RS.EOF
        stHina2.Range(stHina2.Cells(10, 1), stHina2.Cells(10, 14)).Copy ST.Cells(lRow, 1)
        lCol = 3:           ST.Cells(lRow, lCol) = IIf(RS("BANRTM") = 0, "", Left(Format(RS("BANRTM"), "000000"), 2) & ":" & Mid(Format(RS("BANRTM"), "000000"), 3, 2))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BANRTM") = 0, "", Format(RS("BANRTM"), "000000"))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BATNTM") = 0, "", Mid(RS("BATNTM"), 9, 2) & ":" & Mid(RS("BATNTM"), 11, 2))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BATNTM") = 0, "", CStr(RS("BATNTM")))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BAHJCD") = "1", "○", IIf(RS("BAHJCD") = "0", "×", ""))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BAKRTM") = 0, "", Mid(Format(RS("BAKRTM"), "000000"), 9, 2) & ":" & Mid(Format(RS("BAKRTM"), "000000"), 11, 2))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BAKRTM") = 0, "99999999999999", CStr(RS("BAKRTM")))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BAKJWT") = 0, "", Format(RS("BAKJWT"), "0.0"))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = IIf(RS("BASPED") = 0, "", Format(RS("BASPED"), "0.00"))
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = RS("BABIKO")
        lCol = lCol + 1:
        lCol = lCol + 1:    ST.Cells(lRow, lCol) = RS("SYKJ")
        If RS("BAFNCD") = "1" Then
            bEnd = True
            ST.Cells(5, 12) = RS("KSSYKJ")
            ST.Cells(6, 12) = Format(RS("BACHKS"), "00000000")
            ST.Cells(5, 13) = RS("KHSYKJ")
            ST.Cells(6, 13) = Format(RS("BACHKH"), "00000000")
            ST.Cells(5, 14) = RS("KLSYKJ")
            ST.Cells(6, 14) = Format(RS("BACHKL"), "00000000")
            ST.Cells(1, 8) = "生産終了"
        End If
        lRow = lRow + 1
        RS.MoveNext
    Loop
    '投入時間、計量時間順にソートする
    Call subSort

    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing
        
    ST.Cells(9, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set ST = Nothing
End Sub

Private Sub subSort()
    Dim lRow        As Long: lRow = 10
    Dim lRow2       As Long
    Dim lStartRow   As Long: lStartRow = 10
    Dim lMaxRow     As Long: lMaxRow = stList2.Cells(stList2.Rows.Count, 6).End(xlUp).Row
    Dim lCol        As Long
    Dim CntT        As Integer: CntT = 0
    Dim strKey      As String
    Dim ST          As Worksheet: Set ST = stList2
    
    'データが0件だったら処理とばす
    If ST.Cells(10, 6) = "" Then GoTo Exit_
    
    '結合セルをはずす
    ST.Range(ST.Cells(10, 2), ST.Cells(lMaxRow, 14)).UnMerge
    '投入時間、計量時間順で並べ替え
    ST.Range(ST.Cells(10, 2), ST.Cells(lMaxRow, 14)).Sort Key1:=ST.Range("F11"), Order1:=xlAscending, Key2:=ST.Range("I11"), Order2:=xlAscending, Header:=xlNo
    '色すべて消す
    ST.Range(ST.Cells(10, 2), ST.Cells(lMaxRow, 14)).Interior.ColorIndex = xlNone
    '連番消す
    ST.Range(ST.Cells(10, 2), ST.Cells(lMaxRow, 2)).ClearContents                    '投入連番
    '結合しなおす
    Application.DisplayAlerts = False
    Do While Not ST.Cells(lRow, 6) = ""
        If InStr(strKey, ST.Cells(lRow, 6)) = 0 Then
            If Not strKey = "" Then
                strKey = strKey & ","
                '同じ投入時間のデータは結合する
                For lCol = 2 To 7
                    If lCol = 2 Or lCol = 3 Or lCol = 5 Or lCol = 7 Then
                        ST.Range(ST.Cells(lStartRow, lCol), ST.Cells(lRow - 1, lCol)).Merge
                    End If
                Next
            End If
            strKey = strKey & ST.Cells(lRow, 6)
            lStartRow = lRow
            CntT = CntT + 1
        End If
        ST.Cells(lRow, 2) = CStr(CntT)                          '投入連番いれる
        ST.Range(ST.Cells(lRow, 12), ST.Cells(lRow, 13)).Merge  '備考欄マージ
        lRow = lRow + 1
    Loop
    If Not strKey = "" Then
        For lCol = 2 To 7
            If lCol = 2 Or lCol = 3 Or lCol = 5 Or lCol = 7 Then
                ST.Range(ST.Cells(lStartRow, lCol), ST.Cells(lRow - 1, lCol)).Merge
            End If
        Next
    End If
    Application.DisplayAlerts = True

    ST.Range(ST.Cells(10, 2), ST.Cells(lMaxRow, 14)).Borders.LineStyle = xlContinuous
Exit_:
    Set ST = Nothing
End Sub



