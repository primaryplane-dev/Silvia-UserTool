Attribute VB_Name = "Bas_Update"
Option Explicit

Public Sub subUpdate()
    Dim ST          As Object:  Set ST = stList
    Dim CN          As Object:  Set CN = Bas_DbConnection.GetConnection()
    Dim strSQL      As String
    Dim lResult     As Long
    Dim lRow        As Long
    Dim dltFLG      As Boolean: dltFLG = False
    Dim i           As Long
    Dim MaxRow      As Long
    Dim sWorkKTCD   As String
    Dim sEditKTCD   As String
        
    '最大行数取得
    MaxRow = fncGetMaxRow
    
    'ＤＢ接続
    ' DB接続は共通関数で取得済み

    '================================================================
    ' 1. 削除処理 (論理削除)
    '================================================================
    lRow = RW_FR
    Do While Not stWork.Cells(lRow, 1) = ""
        dltFLG = True
        sWorkKTCD = fncGetKTCD(Trim(stWork.Cells(lRow, 1))) '退避データの工程コード
        
        For i = RW_FR To MaxRow
            sEditKTCD = fncGetKTCD(Trim(ST.Cells(i, 1))) '編集データの工程コード
            
            '★Noだけでなく「工程」も一致しているか確認する
            If sWorkKTCD = sEditKTCD And _
               Val(stWork.Cells(lRow, 2)) = Val(StrConv(Trim(ST.Cells(i, 2)), vbNarrow)) Then
                dltFLG = False
                Exit For
            End If
        Next
        
        'シートから消えている場合は削除フラグを立てる
        If dltFLG Then
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBSMF17.SBFP01 SET"
            strSQL = strSQL & "     BFUNTH = TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "    ,BFUTIM = TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & "    ,BFDELT = 'X'"
            strSQL = strSQL & " WHERE BFDELT='' "
            'KEY条件: Noだけでなく工程も指定
            strSQL = strSQL & "   AND BFCVNO=" & Val(stWork.Cells(lRow, 2))
            strSQL = strSQL & "   AND BFKTCD='" & sWorkKTCD & "'" '★追加
            
            CN.Execute strSQL, lResult, &H80
        End If
        lRow = lRow + 1
    Loop
    
    '================================================================
    ' 2. 更新・新規登録処理
    '================================================================
    For lRow = RW_FR To MaxRow
        '必須項目(工程、No、名称)が入っている行のみ対象
        If Trim(ST.Cells(lRow, 1)) <> "" And Val(StrConv(Trim(ST.Cells(lRow, 2)), vbNarrow)) > 0 And Trim(ST.Cells(lRow, 3)) <> "" Then
            
            sEditKTCD = fncGetKTCD(Trim(CStr(ST.Cells(lRow, 1))))
            
            '------------------------------------------------------------
            ' UPDATE (既存データの更新)
            '------------------------------------------------------------
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBSMF17.SBFP01 SET"
            strSQL = strSQL & "     BFUNTH =TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & "    ,BFUTIM =TO_CHAR(current timestamp, 'HH24MISS')"
            '名称更新
            strSQL = strSQL & "    ,BFCVNM = '" & Trim(CStr(ST.Cells(lRow, 3))) & "'"
            '工程コード更新
            strSQL = strSQL & "    ,BFKTCD = '" & sEditKTCD & "'"
            
            strSQL = strSQL & " WHERE BFDELT='' "
            'KEY条件: Noだけでなく工程も指定して、正しい行だけを更新する
            strSQL = strSQL & "   AND BFCVNO=" & Val(StrConv(Trim(ST.Cells(lRow, 2)), vbNarrow))
            strSQL = strSQL & "   AND BFKTCD='" & sEditKTCD & "'" '★追加
            
            CN.Execute strSQL, lResult, &H80
            
            '------------------------------------------------------------
            ' INSERT (該当データが無ければ新規登録)
            '------------------------------------------------------------
            ' 上記UPDATEで「工程」も含めて検索し、ヒットしなければ新規登録となります
            If lResult = 0 Then
                strSQL = ""
                strSQL = strSQL & " INSERT INTO LIBSMF17.SBFP01"
                strSQL = strSQL & "         ("
                strSQL = strSQL & "              BFDELT"
                strSQL = strSQL & "             ,BFCPGM"
                strSQL = strSQL & "             ,BFCNTH"
                strSQL = strSQL & "             ,BFCTIM"
                strSQL = strSQL & "             ,BFKTCD"
                strSQL = strSQL & "             ,BFCVNO"
                strSQL = strSQL & "             ,BFCVNM"
                strSQL = strSQL & "         ) VALUES ("
                strSQL = strSQL & "              ''"
                strSQL = strSQL & "             ,'" & P_PGM & "'"
                strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'YYYYMMDD')"
                strSQL = strSQL & "             ,TO_CHAR(current timestamp, 'HH24MISS')"
                strSQL = strSQL & "             ,'" & sEditKTCD & "'"
                strSQL = strSQL & "             ," & Val(StrConv(Trim(ST.Cells(lRow, 2)), vbNarrow))
                strSQL = strSQL & "             ,'" & Trim(CStr(ST.Cells(lRow, 3))) & "'"
                strSQL = strSQL & "         )"
                
                CN.Execute strSQL, lResult, &H80
            End If
            
        End If
    Next
    
    'DB更新後に最新データでシート再描画
    Call subEditList
    'ＤＢ切断
    CN.Close:   Set CN = Nothing

End Sub

