Option Explicit

' システム管理者情報
Private Const P_SYCD As String = "999999"
Private Const P_SYNM As String = "システム管理者"


' シート定義
Private Property Get stHina() As Worksheet: Set stHina = ThisWorkbook.Worksheets("Hina"): End Property
Private Property Get stList() As Worksheet: Set stList = ThisWorkbook.Worksheets("List"): End Property
Private Property Get stSetting() As Worksheet: Set stSetting = ThisWorkbook.Worksheets("Setting"): End Property

' ==============================================================================
' メイン処理
' ==============================================================================
Public Sub subMain()
   
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Call subBeforeEdit
    Call subEditList
'    Call subEditList2
    Call subAfterEdit
    
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。", vbInformation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

Public Sub subMain2()
    P_DATE = Date
    P_ChkKBN = 2
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Call subBeforeEdit
    Call subEditList
    Call subAfterEdit
    
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' 前処理
Private Sub subBeforeEdit()
    'Application.Calculation = xlCalculationManual ' 必要に応じて
End Sub

' 後処理
Private Sub subAfterEdit()
    'Application.Calculation = xlCalculationAutomatic ' 必要に応じて
End Sub

' 初期化処理
Private Sub subInitialize()
    stHina.Cells.Copy stList.Cells
    stList.Activate
    stList.Range("A1").Select
End Sub

' ==============================================================================
' 一覧表作成処理 (設定シート反映 -> マスタ読込 -> 実績展開)
' ==============================================================================
Public Sub subEditList()
    Dim ST          As Worksheet: Set ST = stList
    Dim SET_WS      As Worksheet: Set SET_WS = stSetting
    Dim RS          As Object
    Dim CN          As Object
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lCol        As Long
    Dim strKey      As String
    Dim strHinoEsc  As String
    Dim strKjnoEsc  As String
    
    ' ---------------------------------------------------------
    ' 1. 初期化 & ヘッダ情報セット
    ' ---------------------------------------------------------
    Call subInitialize ' 雛形コピー (Hina -> List)
    
    ' 雛形をコピーした直後に、Setting2から枠線と図形を上書きコピー
    Call subEditList2
    
    Dim targetCol As Long
    If P_ChkKBN = 2 Then targetCol = 3 Else targetCol = 2

    ' 日付・商品名セット
    ST.Cells(3, 5).Value = Format(P_DATE, "yyyy 年 m 月 d 日")
    ST.Range("J3").Value = P_HINM
    
    ' タイトル・色・名前の設定
    Dim sColor As Long
    If P_ChkKBN = 2 Then
        sColor = RGB(221, 235, 247) ' 薄い青
        ST.Range("Y2").Value = "冷却"
    Else
        sColor = RGB(226, 239, 218) ' 薄い緑
        ST.Range("Y2").Value = "成型"
    End If
    ST.Range("A1").Value = SET_WS.Cells(2, targetCol).Value
    ST.Range("Y2").Interior.Color = sColor
    
    ' 名前の展開 (F34, L34... 7ブロック分)
    Dim sMemberNames As String
    Dim blockNo As Integer
    Dim destCol As Long
    sMemberNames = SET_WS.Cells(3, targetCol).Value
    For blockNo = 0 To 6
        destCol = 6 + (blockNo * 6)
        ST.Cells(ROW_NAME, destCol).Value = sMemberNames
        ST.Cells(ROW_NAME, destCol).WrapText = True
    Next blockNo
    
    ' システム管理用
    ST.Cells(1, CL_SY).Value = ""
    ST.Cells(4, CL_SY).Value = P_SYCD
    ST.Cells(6, CL_SY).Value = P_SYNM

    strHinoEsc = Replace(Trim$(P_HINO & ""), "'", "''")
    strKjnoEsc = Replace(Trim$(P_KJNO & ""), "'", "''")
    
    ' ---------------------------------------------------------
    ' 2. マスタデータ(SBFP01) 読込（BLL経由）
    ' ---------------------------------------------------------
    Set RS = Bas_LogicConveyor.GetConveyorListWithLogic()
    lRow = RW_FR ' データ開始行
    If Not RS Is Nothing Then
        Do While Not RS.EOF
            If lRow > RW_FR Then
                ST.Rows(lRow & ":" & lRow + 1).Insert Shift:=xlDown
                ST.Rows(RW_FR & ":" & RW_FR + 1).Copy Destination:=ST.Rows(lRow)
            End If
            ST.Cells(lRow, 1).Value = RS("No")
            ST.Cells(lRow, 2).Value = RS("名称")
            ST.Cells(lRow, "BC").Value = RS("工程")
            lRow = lRow + 2
            RS.MoveNext
        Loop
        RS.Close
    End If
    ST.Columns("BC").Hidden = True
    
    ' =========================================================
    ' ★修正：SQL0514エラー対策（完全に接続を作り直す）
    ' =========================================================
    If RS.State = 1 Then RS.Close
    If CN.State = 1 Then CN.Close
    Set CN = Nothing
    Set CN = CreateObject("ADODB.Connection")
    CN.CursorLocation = 3
    CN.Open P_ConnectString
    ' =========================================================
    
    ' ---------------------------------------------------------
    ' 3. 実績データ(SBGP01) 読込 & 展開
    ' ---------------------------------------------------------
    strSQL = ""
    strSQL = strSQL & "SELECT BG.* "
    strSQL = strSQL & "      , COALESCE(NULLIF(TS1.TSSYKJ,''), VA1.VASYKJ, '') AS SYKJ "
    strSQL = strSQL & "      , COALESCE(NULLIF(TS2.TSSYKJ,''), VA2.VASYKJ, '') AS KKSYKJ "
    strSQL = strSQL & "  FROM LIBSMF17.SBGP01 AS BG "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS1 ON TS1.TSDELT = '' AND TS1.TSSYCD = BG.BGSGCD AND TS1.TSTTKB = '1' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01   AS VA1 ON VA1.VAKYUK = '' AND VA1.VASYCD = BG.BGSGCD "
    strSQL = strSQL & "  LEFT JOIN LIBSMF17.STSP01 AS TS2 ON TS2.TSDELT = '' AND TS2.TSSYCD = BG.BGCHKK AND TS2.TSTTKB = '4' "
    strSQL = strSQL & "  LEFT JOIN LIBBMF.BVAP01   AS VA2 ON VA2.VAKYUK = '' AND VA2.VASYCD = BG.BGCHKK "
    
    strSQL = strSQL & " WHERE COALESCE(BG.BGDELT,'') = '' "
    strSQL = strSQL & "   AND BG.BGSDAT = " & Val(Format(P_DATE, "yyyymmdd"))
    strSQL = strSQL & "   AND TRIM(BG.BGHINO) = '" & strHinoEsc & "'"
    strSQL = strSQL & "   AND TRIM(BG.BGKJNO) = '" & strKjnoEsc & "'"
    strSQL = strSQL & "   AND COALESCE(BG.BGFNCD,'') <> '1' "
    
    If P_ChkKBN = 1 Then
        strSQL = strSQL & "   AND CAST(BG.BGKTCD AS INT) >= 1 AND CAST(BG.BGKTCD AS INT) <= 20 "
    ElseIf P_ChkKBN = 2 Then
        strSQL = strSQL & "   AND CAST(BG.BGKTCD AS INT) >= 21 "
    End If
    strSQL = strSQL & " ORDER BY BG.BGDATE, BG.BGTIME, BG.BGKTCD, BG.BGCVNO "
    
    RS.Open strSQL, CN, 0, 1
    
    ' データが1件もない場合はメッセージを出して終了
    If RS.EOF Then
        MsgBox "指定された日付・品名の実績データが存在しません。", vbExclamation
    End If
    
    ' ブロック番号(0～6)で列を管理する
    Dim blockIdx As Long
    blockIdx = 0
    strKey = ""
    
    Do While Not RS.EOF
        Dim sCurrentTimeKey As String
        sCurrentTimeKey = RS("BGDATE") & Format(RS("BGTIME"), "000000")
        
        ' 日時が変わったら次のブロック(列)へ移動
        If strKey <> sCurrentTimeKey Then
            ' F(6), L(12), R(18), X(24), AD(30), AJ(36), AP(42) 列目へ
            lCol = 6 + (blockIdx * 6)
            
            ' 枠の最大数(7回)以内ならヘッダ情報を書く
            If blockIdx <= 6 Then
                Dim sTime As String
                sTime = Format(RS("BGTIME"), "000000")
                If Len(sTime) = 6 Then
                    ST.Cells(ROW_TIME, lCol).Value = Format(TimeSerial(Left(sTime, 2), Mid(sTime, 3, 2), Right(sTime, 2)), "h:mm")
                End If
                If Trim$(RS("KKSYKJ") & "") <> "" Then
                    ST.Cells(ROW_NAME, lCol).Value = RS("KKSYKJ")
                Else
                    ST.Cells(ROW_NAME, lCol).Value = RS("SYKJ")
                End If
            End If
            
            blockIdx = blockIdx + 1
            strKey = sCurrentTimeKey
        End If
        
        ' 行の検索 & 結果セット (枠の最大数以内のみ書き込み)
        If blockIdx <= 7 Then
            lRow = fncFindRow(RS("BGKTCD"), CLng(RS("BGCVNO")))
            
            If lRow > 0 Then
                Dim sRes As String
                Select Case CStr(RS("BGCKRT"))
                    Case "0": sRes = "×"
                    Case "1": sRes = "○"
                    Case "2": sRes = "－"
                    Case Else: sRes = RS("BGCKRT")
                End Select
                ST.Cells(lRow, lCol).Value = sRes
            End If
        End If
        
        RS.MoveNext
    Loop

    If RS.State = 1 Then RS.Close

    ' 製造終了データの有無を別判定（通常明細では除外済み）
    strSQL = ""
    strSQL = strSQL & "SELECT 1 AS HIT "
    strSQL = strSQL & "  FROM LIBSMF17.SBGP01 AS BG "
    strSQL = strSQL & " WHERE COALESCE(BG.BGDELT,'') = '' "
    strSQL = strSQL & "   AND BG.BGSDAT = " & Val(Format(P_DATE, "yyyymmdd"))
    strSQL = strSQL & "   AND TRIM(BG.BGHINO) = '" & strHinoEsc & "'"
    strSQL = strSQL & "   AND TRIM(BG.BGKJNO) = '" & strKjnoEsc & "'"
    strSQL = strSQL & "   AND COALESCE(BG.BGFNCD,'') = '1' "
    strSQL = strSQL & " FETCH FIRST 1 ROW ONLY"

    RS.Open strSQL, CN, 0, 1
    If Not RS.EOF Then
        ST.Cells(1, CL_SY + 1).Value = "製造終了"
    End If
    
    ' 終了処理
    If RS.State = 1 Then RS.Close
    If CN.State = 1 Then CN.Close
    Set RS = Nothing
    Set CN = Nothing
End Sub

' ==============================================================================
' レイアウト（セルと図形）の切り替え処理
' ==============================================================================
Private Sub subEditList2()
    Dim SET2_WS As Worksheet: Set SET2_WS = ThisWorkbook.Worksheets("Setting2")
    Dim ST As Worksheet: Set ST = stList
    Dim targetRange As String
    Dim shp As Shape
    Dim i As Long
    Dim pasteArea As Range
    Dim rowCount As Long
    Dim baseCell As Range ' ←追加：図形の高さを計算するための基準セル

    ' 1. コピー元の「セル範囲」と「基準位置」を設定
    If P_ChkKBN = 2 Then ' 冷却
        targetRange = "$A$34:$AV$57" ' （※24行に変更済みの場合は自動で計算）
        Set baseCell = SET2_WS.Range("A34")
    Else ' 成型
        targetRange = "$A$3:$AV$26"
        Set baseCell = SET2_WS.Range("A3")
    End If

    ' コピーする行数を取得
    rowCount = SET2_WS.Range(targetRange).Rows.Count

    ' 2. エラー回避の下準備（貼り付け先の結合解除）
    ST.Rows("5:" & 5 + rowCount - 1).UnMerge

    ' 3. 貼り付け先（5行目）を更地にする
    Set pasteArea = ST.Range("A5").Resize(rowCount, SET2_WS.Range(targetRange).Columns.Count)
    pasteArea.Clear

    ' 4. Setting2シートから、セル範囲を「A5」へ上書きコピー
    ' （※この時、プロパティによって一部の図形が勝手についてくることがある）
    SET2_WS.Range(targetRange).Copy Destination:=ST.Range("A5")

    ' 5. Listシートの図形を「一旦すべて削除」（ボタン類は残す）
    ' 前回の古い図形や、手順4で中途半端についてきた図形をここで処理
    For i = ST.Shapes.Count To 1 Step -1
        Set shp = ST.Shapes(i)
        If shp.Type <> 8 And shp.Type <> 12 Then
            shp.Delete
        End If
    Next i

    ' 6. Setting2から、対象範囲の「図形」だけを個別に確実に取り寄せる
    For Each shp In SET2_WS.Shapes
        ' 図形の左上が対象範囲に入っているか確認
        If Not Intersect(shp.TopLeftCell, SET2_WS.Range(targetRange)) Is Nothing Then
            shp.Copy
            ST.Paste
            
            ' A5セルを基準に、元の位置関係を維持して配置
            With ST.Shapes(ST.Shapes.Count)
                .Top = ST.Range("A5").Top + (shp.Top - baseCell.Top)
                .Left = shp.Left
            End With
        End If
    Next shp

    Application.CutCopyMode = False
    ST.Range("A1").Select
End Sub

' ------------------------------------------------------------------------------
' 行検索関数 (工程CD & コンベアNo で検索)
' ※BC列を参照するように修正
' ------------------------------------------------------------------------------
Private Function fncFindRow(ByVal KTCD As String, ByVal CVNO As Long) As Long
    Dim lRow        As Long
    Dim ST          As Worksheet: Set ST = stList
    Dim lMaxRow     As Long

    fncFindRow = 0
    lRow = RW_FR ' 36
    lMaxRow = 500 ' 探索上限
    
    Do While lRow <= lMaxRow
        ' A列(No)が空なら終了（探索終了）
        If ST.Cells(lRow, 1).Value = "" Then Exit Do
        
        ' ★隠しキー(BC列)と No(A列) でマッチング
        If Val(ST.Cells(lRow, "BC").Value) = Val(KTCD) And Val(ST.Cells(lRow, 1).Value) = CVNO Then
            fncFindRow = lRow
            Exit Do
        End If
        
        lRow = lRow + 2
    Loop
    Set ST = Nothing
End Function




' 終了ボタン用マクロ
Public Sub subCloseBook()
    ' 元のコードの移植
    ' Me（シートモジュール）の代わりに ActiveSheet（現在のシート）を使います
    ActiveSheet.Cells(1, 1).Copy
    Application.CutCopyMode = False
    
    ' 保存せずにブックを閉じる
    ThisWorkbook.Close False
End Sub
