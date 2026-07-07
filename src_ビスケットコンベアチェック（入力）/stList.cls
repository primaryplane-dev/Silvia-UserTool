VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ------------------------------------------------------------
' 1. 条件選択ボタン
' ------------------------------------------------------------
Private Sub cmdSelect_Click()
    If Me.Cells(1, CL_SY) = "1" Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Call subOpenSelect
    If P_Regist Then
        ' 担当者名を表示
        Me.Cells(6, 8) = P_SYNM
        Me.Cells(4, CL_SY) = P_SYCD
        Call subMain
    End If
End Sub

Private Sub subOpenSelect()
    Dim obj As New frmSelect
    obj.Show
    Set obj = Nothing
End Sub

' ------------------------------------------------------------
' 2. 新規登録ボタン
' ------------------------------------------------------------
Private Sub cmdNew_Click()
    If P_SYCD = "" Or P_SYNM = "" Then MsgBox ("担当者を選択してください"): Exit Sub
    If P_DATE = 0 Or (P_HINO = "" And P_KJNO = "") Then MsgBox ("生産日と品名を選択してください"): Exit Sub
    
    Dim lRow As Long, bEnd As Boolean
    
    ' 直近のチェックが完了しているか確認
    lRow = 1: bEnd = True
    Do While Not stWork.Cells(lRow, 1) = ""
        ' タイミング S/M/E の行についてチェック
        If stWork.Cells(lRow, 12) = "S" Or stWork.Cells(lRow, 12) = "M" Or stWork.Cells(lRow, 12) = "E" Then
            ' コンベアには頻度がないため、結果(6)または確認者(9)が空ならNG
            If stWork.Cells(lRow, 6) = "" Or stWork.Cells(lRow, 9) = "" Then
                bEnd = False
                Exit Do
            End If
        End If
        lRow = lRow + 1
    Loop
    
    If Not bEnd Then
        MsgBox ("前のチェックが完了していません"): Exit Sub
    End If
    
    P_DateTimeMode = 1
    Call subOpenDateTime
    If Not P_DateTime_FLG Then Exit Sub
    
    ' stWork にデータを追加
    Call subEditWorkList(Format(P_DateTime, "yyyymmddhhmmss"))
    
    ' 画面再描画
    Call subMain2
    
    ' 未保存フラグON
    Me.Cells(1, CL_SY) = "1"
End Sub
    
Private Sub subOpenDateTime()
    Dim obj As New frmDateTime
    
    ' フォームの表示位置を手動設定モードにする
    obj.StartUpPosition = 0
    
    ' F列30行目のセルの位置を基準にする
    obj.Left = Application.Left + ActiveWindow.Left + ActiveWindow.VisibleRange.Cells(1, 6).Left + 50
    obj.Top = Application.Top + ActiveWindow.Top + 150
    
    obj.Show
    Set obj = Nothing
End Sub

' ------------------------------------------------------------
' 3. 製造終了ボタン
' ------------------------------------------------------------
Private Sub cmdSEnd_Click()
    Dim lRow As Long, bEnd As Boolean
    If Not Me.Cells(1, 7) = "製造中" Then Exit Sub
    
    If stWork.Cells(1, 1) = "" Then MsgBox ("データが未入力です"): Exit Sub
    
    lRow = 1: bEnd = True
    Do While Not stWork.Cells(lRow, 1) = ""
        If stWork.Cells(lRow, 12) = "S" Or stWork.Cells(lRow, 12) = "E" Or stWork.Cells(lRow, 12) = "M" Then
            If stWork.Cells(lRow, 6) = "" Or stWork.Cells(lRow, 9) = "" Then
                bEnd = False: Exit Do
            End If
        End If
        lRow = lRow + 1
    Loop
    If Not bEnd Then
        MsgBox ("未入力の項目があります"): Exit Sub
    End If
    
    If MsgBox("生産を終了するとデータの変更・追加・削除ができなくなりますが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    
    Call subEditWorkList3("", "", "", 4, "1", "")
    Me.Cells(1, 7) = "製造終了"
    Me.Cells(1, CL_SY) = "1"
    cmdNew.Enabled = False
End Sub

' ------------------------------------------------------------
' 4. 担当者変更ボタン
' ------------------------------------------------------------
Private Sub cmdSYCD_Click()
    If Me.Cells(1, CL_SY) = "1" Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Call subOpenSYCD
    stHina.Cells.Copy Me.Cells
    Me.Cells(RW_FR - 3, 1).Select
    Me.Cells(6, 8).Value = P_SYNM
    Me.Cells(4, CL_SY) = P_SYCD
End Sub

Private Sub subOpenSYCD()
    Dim obj As New frmSYCD
    obj.Show
    Set obj = Nothing
End Sub

' ------------------------------------------------------------
' 5. 保存ボタン
' ------------------------------------------------------------
Private Sub cmdSave_Click()
    If P_SYCD = "" Or P_SYNM = "" Then MsgBox ("担当者を選択してください"): Exit Sub
    If P_DATE = 0 Or (P_HINO = "" And P_KJNO = "") Then MsgBox ("生産日と品名を選択してください"): Exit Sub
    If Me.Cells(1, CL_SY) = "" Then Exit Sub
    Call subUpdate
    Call subMain
End Sub

' ------------------------------------------------------------
' 6. 閉じるボタン
' ------------------------------------------------------------
Private Sub cmdEnd_Click()
    If Me.Cells(1, CL_SY) = "1" Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    Application.ThisWorkbook.Close False
End Sub

Private Sub cmdUp_Click()
    If ActiveWindow.ScrollRow < RW_FR Then Exit Sub
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 3
End Sub

Private Sub cmdDown_Click()
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 3
End Sub

' ------------------------------------------------------------
' イベント: セル選択時 (入力・変更・削除)
' ------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' クリック可能範囲の判定
    If Not ((Target.Row >= RW_FR - 3 And Target.Row <= P_BKRow) Or (Target.Row = 6 And 14 <= Target.Column And Target.Column <= 16) Or (Target.Row = 1 And Target.Column = 7)) Then Exit Sub
    If Target.Row = P_SGRow Then Exit Sub
    ' データ列より左、またはシステム列より右は除外
    If Target.Column < 7 Or Target.Column > CL_SY - 1 Then Exit Sub
    ' ヘッダ行以外で、列が初期化（グレーアウト等）されていない箇所は除外
    If Not (Target.Row = 6 Or Target.Row = 1) And (Me.Cells(RW_FR - 3, Target.Column) = ":" Or Me.Cells(RW_FR - 2, Target.Column) = "") Then Exit Sub
    
    If P_SYCD = "" Then MsgBox ("担当者を選択してください"): Exit Sub
    If P_DATE = 0 Or (P_HINM = "" And P_KJNM = "") Then MsgBox ("生産日と品名を選択してください"): Exit Sub
    
    Dim lRow        As Long
    Dim bEnd        As Boolean
    Dim strSYNM     As String
    
    ' 製造終了取消
    If Target.Row = 1 Then
        If Not Me.Cells(Target.Row, Target.Column) = "製造終了" Then Exit Sub
        If fncChkK(3) = "" Then Exit Sub
        If MsgBox("製造終了を取り消します。よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
        Me.Cells(5, 14) = "": Me.Cells(6, 14) = ""
        Me.Cells(5, 15) = "": Me.Cells(6, 15) = ""
        Me.Cells(5, 16) = "": Me.Cells(6, 16) = ""
        Me.Cells(Target.Row, Target.Column) = "製造中"
        Call subEditWorkList3("", "", "", 4, "", "")
        cmdNew.Enabled = True
        
    ' 管理者署名
    ElseIf Target.Row = 6 Then
        If Not Me.Cells(1, 7) = "製造終了" Then Exit Sub
        If Target.Column = 14 Then ' 最終確認者
            If Me.Cells(6, 15) = "" Or Me.Cells(6, 16) = "" Then Exit Sub
            If Me.Cells(Target.Row, Target.Column) = "" Then
                strSYNM = fncChkK(3)
                If strSYNM = "" Then MsgBox ("権限がありません"): Exit Sub
                Me.Cells(Target.Row, Target.Column) = strSYNM
                Me.Cells(Target.Row - 1, Target.Column) = P_SYCD
            End If
        ElseIf Target.Column = 15 Or Target.Column = 16 Then ' 品管・ライン長
            If Me.Cells(Target.Row, Target.Column) = "" Then
                strSYNM = fncChkK(2)
                If strSYNM = "" Then MsgBox ("権限がありません"): Exit Sub
                Me.Cells(Target.Row, Target.Column) = strSYNM
                Me.Cells(Target.Row - 1, Target.Column) = P_SYCD
            End If
        End If
        
    ' 時刻変更・削除
    ElseIf Target.Row = RW_FR - 3 Then
        If Me.Cells(1, 7) = "製造終了" Then Exit Sub
        If Me.Cells(Target.Row, Target.Column).Interior.Color = RGB(166, 166, 166) Then Exit Sub
        If Not (Me.Cells(P_SGRow + 1, Target.Column) = P_SYCD Or Me.Cells(P_SGRow + 1, Target.Column) = "") Then
            MsgBox ("チェックした担当者のみ編集可能です"): Exit Sub
        End If
        P_DateTimeMode = 2
        Call subOpenDateTime
        If Not P_DateTime_FLG Then Exit Sub
        If P_DateTimeMode = 3 Then ' 削除
            Call subEditWorkList2(Format(CDate(stList.Cells(RW_FR - 2, Target.Column)), "yyyymmddhhmmss"), "", 2)
        Else ' 編集
            Call subEditWorkList2(Format(CDate(stList.Cells(RW_FR - 2, Target.Column)), "yyyymmddhhmmss"), Format(P_DateTime, "yyyymmddhhmmss"), 1)
        End If
        Call subMain2
        
    ' 管理者変更(各回)
    ElseIf Target.Row = P_KRRow Then
        If Me.Cells(1, 7) = "製造終了" Then Exit Sub
        If Me.Cells(Target.Row, Target.Column).Interior.Color = RGB(166, 166, 166) Then Exit Sub
        bEnd = True
        lRow = 1
        Do While Not stWork.Cells(lRow, 1) = ""
            If stWork.Cells(lRow, 1) = Format(CDate(stList.Cells(RW_FR - 2, Target.Column)), "yyyymmddhhmmss") Then
                ' 頻度無し: 全て必須チェック
                If stWork.Cells(lRow, 6) = "" Then bEnd = False: Exit Do
            End If
            lRow = lRow + 1
        Loop
        If Not bEnd Then MsgBox ("すべての項目が入力されていません"): GoTo Exit_
        P_SelectKBN = 2
        Call subOpenSelectList
        If Not P_Regist3 Then GoTo Exit_
        Call subEditWorkList3(Format(CDate(stList.Cells(RW_FR - 2, Target.Column)), "yyyymmddhhmmss"), "", "", 3, P_SelCD, P_SelNM)
        Me.Cells(P_KRRow, Target.Column) = P_SelNM
        Me.Cells(P_KRRow + 1, Target.Column) = P_SelCD
    
    ' 備考欄変更
    ElseIf Target.Row = P_BKRow Then
        If Me.Cells(1, 7) = "製造終了" Then Exit Sub
        P_BIKO = Me.Cells(Target.Row, Target.Column)
        Call subOpenBIKO
        If Not P_Regist4 Then GoTo Exit_
        Call subEditWorkList3("", "", "", 2, P_BIKO, "")
        Me.Cells(Target.Row, Target.Column) = P_BIKO
        
    ' チェック結果変更 (データ行)
    Else
        If Me.Cells(1, 7) = "製造終了" Then Exit Sub
        If Me.Cells(Target.Row, Target.Column).Interior.Color = RGB(166, 166, 166) Then Exit Sub
        If Not (Me.Cells(P_SGRow + 1, Target.Column) = P_SYCD Or Me.Cells(P_SGRow + 1, Target.Column) = "") Then
            MsgBox ("チェックした担当者のみ編集可能です"): Exit Sub
        End If
        P_SelectKBN = 1
        Call subOpenSelectList
        If Not P_Regist3 Then GoTo Exit_
        
        ' キー項目更新処理 (C列:工程CD, D列:No)
        Dim sTimeStr As String
        sTimeStr = Format(CDate(stList.Cells(RW_FR - 2, Target.Column)), "yyyymmddhhmmss")
        
        Call subEditWorkList3(sTimeStr, stList.Cells(Target.Row, 3), stList.Cells(Target.Row, 4), 1, P_SelCD, "")
        Me.Cells(Target.Row, Target.Column) = P_SelNM
    End If
    Me.Cells(1, CL_SY) = "1"
Exit_:
    Me.Cells(Target.Row, 5).Select
End Sub

' ------------------------------------------------------------
' ★新規作成時のstWorkへのデータ追加処理
'   (ビスケット備品入力と同じロジックで実装)
' ------------------------------------------------------------
Private Sub subEditWorkList(ByVal dateTime As String)
    Dim lRow    As Long
    Dim lRow2   As Long
    
    ' stWorkの最終行を探す
    lRow = 1
    Do While Not stWork.Cells(lRow, 1) = ""
        lRow = lRow + 1
    Loop
    
    ' stListのデータ行（RW_FR=11行目から）をループしてstWorkに追加
    lRow2 = RW_FR
    
    ' C列(工程CD)がある限りデータを追加
    ' ※P_SGRow(作業者欄)の手前までループ
    Do While lRow2 < P_SGRow
        ' 工程CDが入っている行のみ対象
        If stList.Cells(lRow2, 3) <> "" Then
            stWork.Cells(lRow, 1) = dateTime
            stWork.Cells(lRow, 2) = dateTime
            stWork.Cells(lRow, 3) = stList.Cells(lRow2, 3) ' 工程CD (C列)
            stWork.Cells(lRow, 4) = stList.Cells(lRow2, 4) ' No (D列)
            stWork.Cells(lRow, 5) = ""                     ' 頻度なし
            stWork.Cells(lRow, 7) = Format(P_SYCD, "00000000")
            stWork.Cells(lRow, 8) = P_SYNM
            stWork.Cells(lRow, 11) = stList.Cells(P_BKRow, 7) ' 備考欄の値(デフォルト)
            
            lRow = lRow + 1
        End If
        lRow2 = lRow2 + 1
    Loop
End Sub

' 時刻変更、データ削除
Private Sub subEditWorkList2(ByVal dateTime1 As String, ByVal dateTime2 As String, ByVal mode As Long)
    Dim lRow As Long
    lRow = 1
    Do While Not stWork.Cells(lRow, 1) = ""
        If Not stWork.Cells(lRow, 14) = "X" And stWork.Cells(lRow, 1) = dateTime1 Then
            If mode = 1 Then
                stWork.Cells(lRow, 1) = dateTime2
                stWork.Cells(lRow, 2) = dateTime2
            Else
                stWork.Cells(lRow, 14) = "X"
            End If
        End If
        lRow = lRow + 1
    Loop
End Sub

' チェック項目、備考欄、管理者の更新、製造終了切替
Private Sub subEditWorkList3(ByVal dateTime As String, ByVal KTNO As String, ByVal CKNO As String, ByVal mode As Long, ByVal strVal As String, ByVal strVal2 As String)
    Dim lRow As Long
    lRow = 1
    Do While Not stWork.Cells(lRow, 1) = ""
        If Not stWork.Cells(lRow, 14) = "X" Then
            If mode = 2 Then
                stWork.Cells(lRow, 11) = strVal
            ElseIf mode = 4 Then
                stWork.Cells(lRow, 13) = strVal
            Else
                If stWork.Cells(lRow, 1) = dateTime Then
                    If mode = 1 Then
                        ' 工程CDとNoで一致確認
                        If stWork.Cells(lRow, 3) = KTNO And Val(stWork.Cells(lRow, 4)) = Val(CKNO) Then
                            stWork.Cells(lRow, 6) = strVal
                        End If
                    ElseIf mode = 3 Then
                        stWork.Cells(lRow, 9) = strVal
                        stWork.Cells(lRow, 10) = strVal2
                    End If
                End If
            End If
        End If
        lRow = lRow + 1
    Loop
End Sub

Private Sub subOpenSelectList()
    Dim obj As New frmSelectList
    obj.Show
    Set obj = Nothing
End Sub

Private Sub subOpenBIKO()
    Dim obj As New frmBIKO
    obj.Show
    Set obj = Nothing
End Sub

Private Function fncChkK(ByVal i_KBN As Long) As String
    Dim CN As New ADODB.Connection, RS As New ADODB.Recordset, strSQL As String
    fncChkK = ""
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    strSQL = " SELECT * FROM LIBSMF17.STSP01 WHERE TSDELT = '' AND TSTTKB = '" & i_KBN & "' AND TSSYCD = " & Val(P_SYCD)
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If RS.RecordCount > 0 Then fncChkK = RS("TSSYKJ")
    RS.Close: Set RS = Nothing: CN.Close: Set CN = Nothing
End Function
