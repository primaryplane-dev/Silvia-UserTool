Option Explicit

Public Sub subMain()
    Call subBeforeEdit
    Call subEditList
    Call subEditList2
    Call subAfterEdit
End Sub

Public Sub subMain2()
    Call subBeforeEdit
    Call subEditList2
    Call subAfterEdit
End Sub

Private Sub subInitialize()
    Dim lRow As Long

    stHina.Cells.Copy stList.Cells
    ' 前回実行時の非表示行が残らないように初期化
    stList.Cells.EntireRow.Hidden = False
    ' ヘッダー領域は雛型シートの非表示設定を優先
    For lRow = 1 To RW_FR - 1
        stList.Rows(lRow).Hidden = stHina.Rows(lRow).Hidden
    Next
    ' ビスケット備品入力同様、初期状態の不要行を削除して詰める
    stList.Range("12:16").Delete
    stList.Select
End Sub

' マスタデータ (SBFP01) の読み込みと画面構築
Private Sub subEditList()
    Dim ST          As Worksheet: Set ST = stList
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim startRow    As Long
    Dim strKey      As String
    Dim vKTCD       As Variant
    Dim vCVNM       As Variant
    
    '雛型シート→編集シート
    Call subInitialize
    
    '見出しセット
    ST.Cells(1, 2) = "コンベアチェック表(" & IIf(P_ChkKBN = 1, "成型)", "冷却)")
    ST.Cells(4, 5) = Format(P_DATE, "yyyy年m月d日")
    ST.Cells(4, 8) = IIf(P_KJNO = "", P_HINM, P_KJNO & " " & P_KJNM)
    ST.Cells(6, 8) = P_SYNM
    ST.Cells(4, CL_SY) = Format(P_SYCD, "00000000")
    ST.Cells(1, 7) = "製造中"
    stList.cmdNew.Enabled = True
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    '----------------------------------------------------------------
    ' 1. SBFP01読込 (コンベア名称マスタ)
    '----------------------------------------------------------------
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "   BFKTCD "
    strSQL = strSQL & "  ,BFCVNO "
    strSQL = strSQL & "  ,BFCVNM "
    strSQL = strSQL & "  FROM LIBSMF17.SBFP01 "
    strSQL = strSQL & " WHERE BFDELT = '' "
    
    ' 工程区分絞り込み
    If P_ChkKBN = 1 Then
        strSQL = strSQL & "   AND INT(BFKTCD) >= 1 AND INT(BFKTCD) <= 20 "
    ElseIf P_ChkKBN = 2 Then
        strSQL = strSQL & "   AND INT(BFKTCD) >= 21 "
    End If
    
    strSQL = strSQL & " ORDER BY BFKTCD, BFCVNO "
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    lRow = RW_FR '11行目から開始
    Do While Not RS.EOF
        vKTCD = RS("BFKTCD") & ""
        vCVNM = RS("BFCVNM") & ""
        
        ' 雛型(stHina)のRW_FR行目の書式をコピー
        stHina.Range(stHina.Cells(RW_FR, 1), stHina.Cells(RW_FR, 17)).Copy ST.Cells(lRow, 1)
        
        ' 列セット (ビスケット備品入力のレイアウトに準拠)
        ST.Cells(lRow, 2) = fncGetKTNM(CStr(vKTCD)) '工程名
        ST.Cells(lRow, 3) = vKTCD                   '工程CD(隠し)
        ST.Cells(lRow, 4) = CStr(RS("BFCVNO"))      'No
        ST.Cells(lRow, 5) = vCVNM                   '名称
        ST.Cells(lRow, 6) = ""                      '頻度(SBFP01には無いので空欄)
        
        lRow = lRow + 1
        RS.MoveNext
    Loop
    RS.Close
    
    ' フッター行の位置を決定
    P_SGRow = lRow
    P_KRRow = lRow + 2
    P_BKRow = lRow + 4
    
    ' 作業者、管理者、備考欄を挿入 (雛型の12～16行目をコピー)
    stHina.Range(stHina.Cells(12, 1), stHina.Cells(16, 17)).Copy ST.Cells(P_SGRow, 1)
    ' 雛型側の非表示状態が残るのを防ぐため、挿入範囲は一旦すべて表示
    ST.Rows(P_SGRow & ":" & P_BKRow + 1).Hidden = False
    ST.Rows(P_SGRow + 1).Hidden = True
    ST.Rows(P_KRRow + 1).Hidden = True
    
    ' 工程名のマージ処理
    Application.DisplayAlerts = False
    strKey = "": startRow = RW_FR
    For lRow = RW_FR To P_SGRow - 1
        If Not strKey = ST.Cells(lRow, 3) Then
            If Not strKey = "" Then
                ST.Range(ST.Cells(startRow, 2), ST.Cells(lRow - 1, 2)).Merge
            End If
            startRow = lRow
            strKey = ST.Cells(lRow, 3)
        End If
    Next
    If Not strKey = "" Then
        ST.Range(ST.Cells(startRow, 2), ST.Cells(P_SGRow - 1, 2)).Merge
    End If
    Application.DisplayAlerts = True
    
    ' 罫線
    ST.Range(ST.Cells(RW_FR, 2), ST.Cells(P_BKRow, 16)).Borders.LineStyle = xlContinuous
    
    ' stWork初期化
    stWork.Cells.ClearContents
    
    '----------------------------------------------------------------
    ' 2. SBGP01読込 (実績データ)
    '----------------------------------------------------------------
    strSQL = ""
    strSQL = strSQL & "SELECT BG.* "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS1.TSSYKJ,'')='' THEN VA1.VASYKJ ELSE COALESCE(TS1.TSSYKJ,'') END AS SYKJ "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS2.TSSYKJ,'')='' THEN VA2.VASYKJ ELSE COALESCE(TS2.TSSYKJ,'') END AS KKSYKJ "
    strSQL = strSQL & "  FROM LIBSMF17.SBGP01 AS BG "
    
    ' 結合条件はビスケット備品入力と同じロジックを使用
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 ON TS1.TSDELT = '' AND TS1.TSSYCD = BG.BGSGCD AND TS1.TSTTKB = '1' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA1 ON VA1.VAKYUK = '' AND VA1.VASYCD = BG.BGSGCD "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 ON TS2.TSDELT = '' AND TS2.TSSYCD = BG.BGCHKK AND TS2.TSTTKB = '4' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA2 ON VA2.VAKYUK = '' AND VA2.VASYCD = BG.BGCHKK "
    
    strSQL = strSQL & " WHERE BG.BGDELT = '' "
    strSQL = strSQL & "   AND BG.BGSDAT = " & Val(Format(P_DATE, "yyyymmdd"))
    strSQL = strSQL & "   AND BG.BGHINO = " & Val(P_HINO)
    strSQL = strSQL & "   AND BG.BGKJNO = '" & P_KJNO & "'"
    
    If P_ChkKBN = 1 Then
        strSQL = strSQL & "   AND INT(BG.BGKTCD) >= 1 AND INT(BG.BGKTCD) <= 20 "
    ElseIf P_ChkKBN = 2 Then
        strSQL = strSQL & "   AND INT(BG.BGKTCD) >= 21 "
    End If
    
    strSQL = strSQL & " ORDER BY BG.BGDATE, BG.BGTIME "
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lRow = 1
    Do While Not RS.EOF
        ' stWorkへの展開 (ビスケット備品入力のレイアウトに合わせる)
        stWork.Cells(lRow, 1) = RS("BGDATE") & Format(RS("BGTIME"), "000000")
        stWork.Cells(lRow, 2) = RS("BGDATE") & Format(RS("BGTIME"), "000000")
        stWork.Cells(lRow, 3) = RS("BGKTCD")
        stWork.Cells(lRow, 4) = RS("BGCVNO")
        stWork.Cells(lRow, 5) = ""           '頻度なし
        stWork.Cells(lRow, 6) = RS("BGCKRT")
        stWork.Cells(lRow, 7) = RS("BGSGCD")
        stWork.Cells(lRow, 8) = RS("SYKJ")
        stWork.Cells(lRow, 9) = RS("BGCHKK")
        stWork.Cells(lRow, 10) = RS("KKSYKJ")
        stWork.Cells(lRow, 11) = RS("BGBIKO")
        stWork.Cells(lRow, 12) = RS("BGTMCD")
        stWork.Cells(lRow, 13) = RS("BGFNCD")
        stWork.Cells(lRow, 14) = ""
        lRow = lRow + 1
        RS.MoveNext
    Loop
    RS.Close
    
    '----------------------------------------------------------------
    ' 3. 管理者(ライン長・品管・最終確認者)取得
    '----------------------------------------------------------------
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT "
    strSQL = strSQL & "       BGCHKL "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS1.TSSYKJ,'')='' THEN VA1.VASYKJ ELSE COALESCE(TS1.TSSYKJ,'') END AS KLSYKJ "
    strSQL = strSQL & "     , BGCHKH "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS2.TSSYKJ,'')='' THEN VA2.VASYKJ ELSE COALESCE(TS2.TSSYKJ,'') END AS KHSYKJ "
    strSQL = strSQL & "     , BGCHKS "
    strSQL = strSQL & "     , CASE WHEN COALESCE(TS3.TSSYKJ,'')='' THEN VA3.VASYKJ ELSE COALESCE(TS3.TSSYKJ,'') END AS KSSYKJ "
    strSQL = strSQL & "  FROM LIBSMF17.SBGP01 "
    
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 ON TS1.TSDELT = '' AND TS1.TSSYCD = BGCHKL AND TS1.TSTTKB = '2'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA1 ON VA1.VAKYUK = '' AND VA1.VASYCD = BGCHKL"
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 ON TS2.TSDELT = '' AND TS2.TSSYCD = BGCHKH AND TS2.TSTTKB = '2'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA2 ON VA2.VAKYUK = '' AND VA2.VASYCD = BGCHKH "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS3 ON TS3.TSDELT = '' AND TS3.TSSYCD = BGCHKS AND TS3.TSTTKB = '3'"
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01 AS VA3 ON VA3.VAKYUK = '' AND VA3.VASYCD = BGCHKS "
    
    strSQL = strSQL & " WHERE BGDELT='' "
    strSQL = strSQL & "   AND BGSDAT=" & Val(Format(P_DATE, "yyyymmdd"))
    strSQL = strSQL & "   AND BGHINO=" & Val(P_HINO)
    strSQL = strSQL & "   AND BGKJNO='" & P_KJNO & "'"
    strSQL = strSQL & "   AND BGFNCD='1' " '製造終了データ
    
    If P_ChkKBN = 1 Then
        strSQL = strSQL & "   AND INT(BGKTCD) >= 1 AND INT(BGKTCD) <= 20 "
    ElseIf P_ChkKBN = 2 Then
        strSQL = strSQL & "   AND INT(BGKTCD) >= 21 "
    End If
    strSQL = strSQL & "   AND (BGCHKL<>'' OR BGCHKH<>'' OR BGCHKS<>'')"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    If RS.RecordCount > 0 Then
        ST.Cells(5, 14) = RS("BGCHKS")
        ST.Cells(6, 14) = RS("KSSYKJ")
        ST.Cells(5, 15) = RS("BGCHKH")
        ST.Cells(6, 15) = RS("KHSYKJ")
        ST.Cells(5, 16) = RS("BGCHKL")
        ST.Cells(6, 16) = RS("KLSYKJ")
    End If
    
    If RS.State = 1 Then RS.Close
    Set RS = Nothing
    CN.Close: Set CN = Nothing
    Set ST = Nothing
End Sub

Private Sub subEditList2()
    Dim ST              As Worksheet: Set ST = stList
    Dim lRow            As Long
    Dim lRow2           As Long
    Dim lCol            As Long
    Dim strKey          As String
    Dim Cnt             As Long
    
    'ソート、タイミングコード振り
    lRow = 1: strKey = "": Cnt = 0
    If Not stWork.Cells(1, 1) = "" Then
        stWork.Range(stWork.Cells(1, 1), stWork.Cells(5000, 14)).Sort Key1:=stWork.Range("A1"), Order1:=xlAscending
        Do While Not stWork.Cells(lRow, 1) = ""
            If Not strKey = stWork.Cells(lRow, 1) & stWork.Cells(lRow, 14) Then
                If Not stWork.Cells(lRow, 14) = "X" Then
                    strKey = stWork.Cells(lRow, 1) & stWork.Cells(lRow, 14)
                    Cnt = Cnt + 1
                End If
            End If
            If Not stWork.Cells(lRow, 14) = "X" Then
                stWork.Cells(lRow, 12) = CStr(Cnt)
            Else
                stWork.Cells(lRow, 12) = ""
            End If
            lRow = lRow + 1
        Loop
    End If
    
    'データエリア初期化 (RW_FRより下のデータ列をクリア)
    ST.Range(ST.Cells(RW_FR - 3, 7), ST.Cells(RW_FR - 3, 16)) = ":"
    ST.Range(ST.Cells(RW_FR - 3, 7), ST.Cells(RW_FR - 3, 16)).Interior.Color = RGB(242, 242, 242)
    ST.Range(ST.Cells(RW_FR - 2, 7), ST.Cells(RW_FR - 1, 16)).ClearContents
    ST.Range(ST.Cells(RW_FR - 2, 7), ST.Cells(RW_FR - 1, 16)).Interior.ColorIndex = xlNone
    ST.Range(ST.Cells(RW_FR, 7), ST.Cells(P_KRRow, 16)).ClearContents
    ST.Cells(P_BKRow, 7) = ""
    ST.Range(ST.Cells(RW_FR, 7), ST.Cells(P_KRRow, 16)).Interior.ColorIndex = xlNone
    
    strKey = "": lRow = 1: lCol = 6 'G列(7)の1つ前=6からスタート
    
    Do While Not stWork.Cells(lRow, 1) = ""
        ' S/M/E
        If Cnt > 0 Then
             If Val(stWork.Cells(lRow, 12)) = 1 Then stWork.Cells(lRow, 12) = "S"
             If Val(stWork.Cells(lRow, 12)) > 1 And Val(stWork.Cells(lRow, 12)) < Cnt Then stWork.Cells(lRow, 12) = "M"
        End If
        If Cnt > 1 Then
             If Val(stWork.Cells(lRow, 12)) = Cnt Then stWork.Cells(lRow, 12) = "E"
        End If
        
        ' 列切り替え
        If Not strKey = stWork.Cells(lRow, 1) & stWork.Cells(lRow, 14) Then
            lCol = lCol + 1
            ST.Cells(RW_FR - 3, lCol) = Format(CDate(Mid(stWork.Cells(lRow, 1), 9, 2) & ":" & Mid(stWork.Cells(lRow, 1), 11, 2) & ":" & Right(stWork.Cells(lRow, 1), 2)), "H:MM")
            ST.Cells(RW_FR - 2, lCol) = Format(Left(stWork.Cells(lRow, 1), 8), "0000/00/00") & " " & Format(Right(stWork.Cells(lRow, 1), 6), "00:00:00")
            ST.Cells(RW_FR - 1, lCol) = Format(Left(stWork.Cells(lRow, 2), 8), "0000/00/00") & " " & Format(Right(stWork.Cells(lRow, 2), 6), "00:00:00")
            strKey = stWork.Cells(lRow, 1) & stWork.Cells(lRow, 14)
        End If
        
        ' 行検索 (工程+No)
        lRow2 = fncFindRow(stWork.Cells(lRow, 3), CLng(stWork.Cells(lRow, 4)))
        If lRow2 > 0 Then
            ' 結果
            If stWork.Cells(lRow, 6) = "0" Then ST.Cells(lRow2, lCol) = "×"
            If stWork.Cells(lRow, 6) = "1" Then ST.Cells(lRow2, lCol) = "○"
            If stWork.Cells(lRow, 6) = "2" Then ST.Cells(lRow2, lCol) = "－"
            
            ' 担当者・確認者
            ST.Cells(P_SGRow, lCol) = stWork.Cells(lRow, 8)
            ST.Cells(P_SGRow + 1, lCol) = Format(stWork.Cells(lRow, 7), "00000000")
            ST.Cells(P_KRRow, lCol) = stWork.Cells(lRow, 10)
            ST.Cells(P_KRRow + 1, lCol) = Format(stWork.Cells(lRow, 9), "00000000")
            ST.Cells(P_BKRow, 7) = stWork.Cells(lRow, 11)
        End If
        
        ' 削除行の色付け
        If stWork.Cells(lRow, 14) = "X" Then
            ST.Range(ST.Cells(RW_FR - 3, lCol), ST.Cells(P_KRRow, lCol)).Interior.Color = RGB(166, 166, 166)
        End If
        
        ' 製造終了
        If stWork.Cells(lRow, 13) = "1" Then
            ST.Cells(1, 7) = "製造終了"
            stList.cmdNew.Enabled = False
        End If
        
        lRow = lRow + 1
    Loop
    
    
    ' 印刷範囲の設定
    ST.PageSetup.PrintArea = "$A$1:$R$" & P_BKRow
    
    Set ST = Nothing
    
    
End Sub

Private Function fncFindRow(ByVal KTCD As String, ByVal CVNO As Long) As Long
    Dim lRow        As Long
    Dim ST          As Worksheet: Set ST = stList
    Dim strSheetKT  As String

    fncFindRow = 0
    lRow = RW_FR
    Do While Not ST.Cells(lRow, 3) = ""
        strSheetKT = ST.Cells(lRow, 3) & ""
        If strSheetKT = KTCD And Val(ST.Cells(lRow, 4)) = CVNO Then
            fncFindRow = lRow
            Exit Do
        End If
        lRow = lRow + 1
    Loop
    Set ST = Nothing
End Function

' 工程CD→名称変換 (簡易版)
Public Function fncGetKTNM(ByVal KTCD As String) As String
    Select Case KTCD
        Case "1": fncGetKTNM = "成型"
        Case "21": fncGetKTNM = "冷却"
        Case Else: fncGetKTNM = ""
    End Select
End Function
