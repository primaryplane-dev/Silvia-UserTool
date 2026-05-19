Option Explicit

Public Const RW_FR = 4

Public Sub subDisplayList()
    Call subBeforeEdit
    Call subEditList
    Call subAfterEdit
End Sub

Private Sub subInitialize()
    stHina.Cells.Copy stList.Cells
    stList.Select
End Sub

Private Sub subEditList()
    Dim ST          As Worksheet: Set ST = stList
    Dim CN          As ADODB.Connection
    Dim RS          As ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lCol        As Long
    Dim lNo         As Long
    Dim lLen        As Long
    Dim strWK       As String
    Dim sSp()       As String
    Dim i           As Long
    Dim strKey      As String
    Dim strFROM     As String
    Dim strTO       As String
    
    If P_DATEF = 0 Then
        strFROM = "0"
    Else
        strFROM = Format(P_DATEF, "yyyymmdd")
    End If
    If P_DATET = 0 Then
        strTO = Format(Date, "yyyymmdd")
    Else
        strTO = Format(P_DATET, "yyyymmdd")
    End If
    '雛型シート→編集シート
    Call subInitialize

    'ＤＢ接続
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    'SBAP01:ビスケット分割重量チェック
    strSQL = ""
    strSQL = strSQL & " SELECT  DISTINCT "
    strSQL = strSQL & "         BADATE AS SDATE "   '日付
    strSQL = strSQL & "        ,'ビスケット分割重量' AS KBN "   'チェック表区分
    strSQL = strSQL & "        ,CASE WHEN BAKJNM = '' THEN TRIM(BAHINM) ELSE TRIM(BAHINM) || ' : ' || BAKJNM END AS HINM "   '製品名
    strSQL = strSQL & "        ,COALESCE(BACHKL,0) AS CHKL "    'ライン長
    strSQL = strSQL & "        ,COALESCE(BACHKH,0) AS CHKH "    '品管
    strSQL = strSQL & "        ,COALESCE(BACHKS,0) AS CHKS "    '最終確認者
    strSQL = strSQL & "        ,BAHINO AS HINO "    '品番
    strSQL = strSQL & "        ,BAKJNO AS KJNO "    '生地番
    strSQL = strSQL & "        ,0      AS SHBU "    'アイテム種類
    strSQL = strSQL & "        ,0      AS KTCD "    '工程コード（SBAP01は常に0）
    strSQL = strSQL & "    FROM LIBSMF17.SBAP01 "
    strSQL = strSQL & "   WHERE BADELT = ''"
    strSQL = strSQL & "     AND BADATE BETWEEN " & strFROM & " AND " & strTO
    strSQL = strSQL & "     AND (COALESCE(BACHKL,0) = 0 OR COALESCE(BACHKH,0) = 0 OR COALESCE(BACHKS,0) = 0) "
    strSQL = strSQL & "     AND COALESCE(BAFNCD,'') = '1' "
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & " SELECT  DISTINCT"
    strSQL = strSQL & "         BESDAT AS SDATE "   '日付
    strSQL = strSQL & "        ,'備品チェック表' AS KBN "       'チェック表区分
    strSQL = strSQL & "        ,CASE WHEN BEKJNM = '' THEN TRIM(BEHINM) ELSE TRIM(BEHINM) || ' : ' || BEKJNM END AS HINM "    '製品名
    strSQL = strSQL & "        ,COALESCE(BECHKL,0) AS CHKL "    'ライン長
    strSQL = strSQL & "        ,COALESCE(BECHKH,0) AS CHKH "    '品管
    strSQL = strSQL & "        ,COALESCE(BECHKS,0) AS CHKS "    '最終確認者
    strSQL = strSQL & "        ,BEHINO AS HINO "    '品番
    strSQL = strSQL & "        ,BEKJNO AS KJNO "    '生地番
    strSQL = strSQL & "        ,BESHBU AS SHBU "    'アイテム種類
    strSQL = strSQL & "        ,CASE WHEN BEKTCD <= 20 THEN 1 WHEN BEKTCD >= 21 THEN 2 END AS KTCD "    '工程コード（SBEP01）
    strSQL = strSQL & "    FROM LIBSMF17.SBEP01 "
    strSQL = strSQL & "   WHERE BEDELT = ''"
    strSQL = strSQL & "     AND BESDAT BETWEEN " & strFROM & " AND " & strTO
    strSQL = strSQL & "     AND (COALESCE(BECHKL,0) = 0 OR COALESCE(BECHKH,0) = 0 OR COALESCE(BECHKS,0) = 0) "
    strSQL = strSQL & "     AND COALESCE(BEFNCD,'') = '1' "
    strSQL = strSQL & "     AND NOT COALESCE(BESHBU,0) = 0 "
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & " SELECT  DISTINCT"
    strSQL = strSQL & "         BGSDAT AS SDATE "   '日付
    strSQL = strSQL & "        ,'コンベアチェック表' AS KBN "        'チェック表区分
    strSQL = strSQL & "        ,CASE WHEN BGKJNM = '' THEN TRIM(BGHINM) ELSE TRIM(BGHINM) || ' : ' || BGKJNM END AS HINM "    '製品名
    strSQL = strSQL & "        ,COALESCE(BGCHKL,0) AS CHKL "    'ライン長
    strSQL = strSQL & "        ,COALESCE(BGCHKH,0) AS CHKH "    '品管
    strSQL = strSQL & "        ,COALESCE(BGCHKS,0) AS CHKS "    '最終確認者
    strSQL = strSQL & "        ,BGHINO AS HINO "    '品番
    strSQL = strSQL & "        ,BGKJNO AS KJNO "    '生地番
    strSQL = strSQL & "        ,0      AS SHBU "    'アイテム種類
    strSQL = strSQL & "        ,BGKTCD AS KTCD "    '工程コード（SBGP01のみBGKTCD）
    strSQL = strSQL & "    FROM LIBSMF17.SBGP01 "
    strSQL = strSQL & "   WHERE BGDELT = ''"
    strSQL = strSQL & "     AND BGSDAT BETWEEN " & strFROM & " AND " & strTO
    strSQL = strSQL & "     AND (COALESCE(BGCHKL,0) = 0 OR COALESCE(BGCHKH,0) = 0 OR COALESCE(BGCHKS,0) = 0) "
    strSQL = strSQL & "     AND COALESCE(BGFNCD,'') = '1' "
    strSQL = strSQL & "   ORDER BY SDATE, KBN "
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    lRow = RW_FR: strKey = ""
    Do While Not RS.EOF
        ' Debug出力
        Debug.Print "lRow=" & lRow & ", SDATE=" & RS("SDATE") & ", KBN=" & RS("KBN") & ", HINM=" & RS("HINM") & _
            ", CHKL=" & RS("CHKL") & ", CHKH=" & RS("CHKH") & ", CHKS=" & RS("CHKS") & ", HINO=" & RS("HINO") & _
            ", KJNO=" & RS("KJNO") & ", SHBU=" & RS("SHBU") & ", KTCD=" & RS("KTCD")
        ' BGCVNOのNullチェック例（必要な場所に応じて挿入してください）
        'Dim tmpCVNO As Variant
        'tmpCVNO = RS("BGCVNO")
        'If IsNull(tmpCVNO) Or tmpCVNO = "" Then
        '    lRow = 0
        'Else
        '    lRow = fncFindRow(ST, CLng(tmpCVNO), lKTRow)
        'End If
        lCol = 1: ST.Cells(lRow, lCol) = Left(RS("SDATE"), 4) & "/" & Mid(RS("SDATE"), 5, 2) & "/" & Right(RS("SDATE"), 2)
        lCol = lCol + 1: ST.Cells(lRow, lCol) = fncGetSHNM(CStr(RS("SHBU"))) & RS("KBN") & IIf(RS("KBN") = "備品チェック表", "(" & fncGetKTNM2(CStr(RS("KTCD"))) & ")", "")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = RS("HINM")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = IIf(RS("CHKL") > 0, "○", "×")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = IIf(RS("CHKH") > 0, "○", "×")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = IIf(RS("CHKS") > 0, "○", "×")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = RS("HINO")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = RS("KJNO")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = RS("SHBU")
        lCol = lCol + 1: ST.Cells(lRow, lCol) = RS("KTCD")
        lRow = lRow + 1
        RS.MoveNext
    Loop
    
    '罫線を引く
    ST.Range(ST.Cells(RW_FR, 1), ST.Cells(lRow - 1, 10)).Borders.LineStyle = xlContinuous
    
    'DB切断
    RS.Close:
    Set RS = Nothing
    CN.Close: Set CN = Nothing
    
    ST.Cells(RW_FR - 1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set ST = Nothing
    
End Sub

Public Sub subMain4()
    ' 必要な処理をここに記述
End Sub

Public Sub subMain3()
    ' 必要な処理をここに記述
End Sub