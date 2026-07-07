Attribute VB_Name = "Bas_List4"
Option Explicit

Public Sub subMain4()
    Call subBeforeEdit
    Call subEditList4
    Call subAfterEdit
End Sub

Private Sub subInitialize()
    stHina4.Cells.Copy stList4.Cells
    stList4.Range("13:19").Delete
    stList4.Select
End Sub

Private Sub subEditList4()
    Dim ST          As Worksheet: Set ST = stList4
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim startRow    As Long
    Dim lCol        As Long
    Dim lCol2       As Long
    Dim bEnd        As Boolean
    Dim strKey      As String
    Dim Cnt         As Integer
    Dim lEndRow     As Long
    Dim lKTRow      As Long
    Dim lSGRow      As Long
    Dim lKRRow      As Long
    Dim lBKRow      As Long
    
    '雛型シート→編集シート
    Call subInitialize
    
    '見だし
    ST.Cells(5, 3) = Format(P_DATE, "yyyy年m月d日")
    ST.Cells(5, 5) = P_HINM
    ST.Cells(1, 4) = "製造終了"
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    'SBFP01:ビスケットコンベアチェックマスタ
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "   BFCVNO "
    strSQL = strSQL & "  ,BFCVNM "
    strSQL = strSQL & "  FROM LIBSMF17.SBFP01 "
    strSQL = strSQL & " WHERE BFDELT = '' "
    strSQL = strSQL & " ORDER BY BFCVNO "
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    lRow = 12
    Do While Not RS.EOF
        stHina4.Range(stHina4.Cells(12, 1), stHina4.Cells(11, 14)).Copy ST.Cells(lRow, 1)
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
    stHina4.Range(stHina4.Cells(13, 1), stHina4.Cells(19, 14)).Copy ST.Cells(lKTRow, 1)
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
    strSQL = strSQL & " ORDER BY BG.BGDATE, BG.BGTIME "
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lCol = 3: strKey = ""
    Do While Not RS.EOF
        If Not strKey = RS("BGDATE") & Format(RS("BGTIME"), "000000") Then
            lCol = lCol + 1
            ST.Cells(9, lCol) = Format(CDate(Format(RS("BGTIME"), "00:00:00")), "H:MM")
            ST.Cells(10, lCol) = Format(RS("BGDATE"), "0000/00/00") & " " & Format(RS("BGTIME"), "00:00:00")
            ST.Cells(11, lCol) = Format(RS("BGDATE"), "0000/00/00") & " " & Format(RS("BGTIME"), "00:00:00")
            strKey = RS("BGDATE") & Format(RS("BGTIME"), "000000")
        End If
        lRow = fncFindRow(CLng(RS("BGCVNO")), lKTRow)
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
    strSQL = strSQL & "   AND BGFNCD='1' "
    strSQL = strSQL & "   AND (BGCHKL<>'' OR BGCHKH<>'' OR BGCHKS<>'')"

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
    
    Set ST = Nothing
End Sub

Private Function fncFindRow(ByVal CVNO As Long, ByVal lKTRow As Long) As Long
    Dim lRow        As Long
    Dim ST          As Worksheet: Set ST = stList4

    fncFindRow = 0
    If CVNO >= 901 Then
        fncFindRow = lKTRow + CVNO - 901
    Else
        lRow = 12
        Do While Not ST.Cells(lRow, 2) = ""
            If CLng(ST.Cells(lRow, 2)) = CVNO Then
                fncFindRow = lRow
                Exit Do
            End If
            lRow = lRow + 1
        Loop
    End If
    Set ST = Nothing

End Function

Private Function fncGetKigo(ByVal i_KigoCD As String) As String
    fncGetKigo = ""
    Select Case i_KigoCD
    Case "1": fncGetKigo = "△"
    Case "2": fncGetKigo = "◇"
    Case "3": fncGetKigo = "●"
    Case "4": fncGetKigo = "×"
    Case "5": fncGetKigo = "○"
    Case "0": fncGetKigo = "－"
    End Select
End Function
