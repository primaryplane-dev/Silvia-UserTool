VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''
' * @file stList.cls
' * @brief チェックリスト管理クラス (UI/ロジック層)
' * @note チェックリストの操作・データ編集・画面連携を担当
'''
Option Explicit


'/**
' * @brief 条件選択ボタン押下時の処理
' * @note 担当者選択フォームを開き、選択結果を反映する
' * @return なし
' */
Private Sub cmdSelect_Click()
    On Error GoTo ErrorHandler

    If Me.Cells(1, CL_SY) = "1" Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel, Bas_Configuration.SYSTEM_NAME) = vbCancel Then Exit Sub
    End If
    Call subOpenSelect
    If P_Regist Then
        Me.Cells(6, 8) = P_SYNM
        Me.Cells(4, CL_SY) = P_SYCD
        Call subMain
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.cmdSelect_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("条件選択処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 条件選択フォームを開く
' * @note frmSelectフォームをモーダルで表示
' * @return なし
' */
Private Sub subOpenSelect()
    On Error GoTo ErrorHandler

    Dim obj As New frmSelect
    obj.Show
    Set obj = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.subOpenSelect: " & Err.Number & " - " & Err.Description
    Call MsgBox("条件選択フォーム呼び出しでエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub


'/**
' * @brief 新規登録ボタン押下時の処理
' * @note チェックリストの新規データ追加。直前チェック未完了時は警告。
' * @return なし
' */
Private Sub cmdNew_Click()
    On Error GoTo ErrorHandler

    If P_SYCD = "" Or P_SYNM = "" Then Call MsgBox("担当者を選択してください", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    If P_DATE = 0 Or (P_HINO = "" And P_KJNO = "") Then Call MsgBox("生産日と品名を選択してください", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    
    Dim lRow As Long, bEnd As Boolean
    lRow = 1: bEnd = True
    Do While Not stWork.Cells(lRow, 1) = ""
        If stWork.Cells(lRow, 12) = "S" Or stWork.Cells(lRow, 12) = "M" Or stWork.Cells(lRow, 12) = "E" Then
            If stWork.Cells(lRow, 6) = "" Or stWork.Cells(lRow, 9) = "" Then
                bEnd = False
                Exit Do
            End If
        End If
        lRow = lRow + 1
    Loop
    
    If Not bEnd Then
        Call MsgBox("前のチェックが完了していません", vbExclamation, Bas_Configuration.SYSTEM_NAME): Exit Sub
    End If
    P_DateTimeMode = 1
    Call subOpenDateTime
    If Not P_DateTime_FLG Then Exit Sub
    Call subEditWorkList(Format(P_DateTime, "yyyymmddhhmmss"))
    Call subMain2
    Me.Cells(1, CL_SY) = "1"

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.cmdNew_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("新規登録処理でエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub
    
'/**
' * @brief 日時入力フォームを開く
' * @note frmDateTimeフォームを表示し、日時入力を受け付ける
' * @return なし
' */
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

'/**
' * @brief 製造終了ボタン押下時の処理
' * @note チェックリストの製造終了処理。未入力項目があれば警告。
' * @return なし
' */
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

'/**
' * @brief 担当者変更ボタン押下時の処理
' * @note 担当者選択フォームを開き、担当者情報を反映
' * @return なし
' */
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

'/**
' * @brief 担当者選択フォームを開く
' * @note frmSYCDフォームをモーダルで表示
' * @return なし
' */
Private Sub subOpenSYCD()
    Dim obj As New frmSYCD
    obj.Show
    Set obj = Nothing
End Sub

'/**
' * @brief 保存ボタン押下時の処理
' * @note 入力値検証後、チェックリストを保存
' * @return なし
' */
Private Sub cmdSave_Click()
    If P_SYCD = "" Or P_SYNM = "" Then MsgBox ("担当者を選択してください"): Exit Sub
    If P_DATE = 0 Or (P_HINO = "" And P_KJNO = "") Then MsgBox ("生産日と品名を選択してください"): Exit Sub
    If Me.Cells(1, CL_SY) = "" Then Exit Sub
    Call subUpdate
    Call subMain
End Sub

'/**
' * @brief 閉じるボタン押下時の処理
' * @note 保存確認後、ブックを閉じる
' * @return なし
' */
Private Sub cmdEnd_Click()
    If Me.Cells(1, CL_SY) = "1" Then
        If MsgBox("保存されていませんが、よろしいですか？", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    Application.ThisWorkbook.Close False
End Sub

'/**
' * @brief 上スクロールボタン押下時の処理
' * @note 表示行を3行上にスクロール
' * @return なし
' */
Private Sub cmdUp_Click()
    If ActiveWindow.ScrollRow < RW_FR Then Exit Sub
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 3
End Sub

'/**
' * @brief 下スクロールボタン押下時の処理
' * @note 表示行を3行下にスクロール
' * @return なし
' */
Private Sub cmdDown_Click()
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 3
End Sub

'/**
' * @brief セル選択時のイベント
' * @param Target 選択セル範囲
' * @note 入力・編集・削除・管理者署名・備考編集等の分岐処理
' * @return なし
' */
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

'/**
' * @brief stWorkへのデータ追加処理
' * @param dateTime 追加する日時（yyyymmddhhmmss）
' * @note 新規作成時にstListの内容をstWorkへ転記
' * @return なし
' */
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
'/**
' * @brief 時刻変更・データ削除処理
' * @param dateTime1 変更・削除対象日時
' * @param dateTime2 変更後日時（mode=1時のみ）
' * @param mode 1=時刻変更, 2=削除
' * @note stWorkの該当行を更新または削除
' * @return なし
' */
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
'/**
' * @brief チェック項目・備考欄・管理者の更新、製造終了切替
' * @param dateTime 対象日時
' * @param KTNO 工程CD
' * @param CKNO No
' * @param mode 1=チェック項目, 2=備考, 3=管理者, 4=製造終了
' * @param strVal 更新値1
' * @param strVal2 更新値2
' * @note stWorkの該当行を更新
' * @return なし
' */
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

'/**
' * @brief 選択リストフォームを開く
' * @note frmSelectListフォームをモーダルで表示
' * @return なし
' */
Private Sub subOpenSelectList()
    Dim obj As New frmSelectList
    obj.Show
    Set obj = Nothing
End Sub

'/**
' * @brief 備考入力フォームを開く
' * @note frmBIKOフォームをモーダルで表示
' * @return なし
' */
Private Sub subOpenBIKO()
    Dim obj As New frmBIKO
    obj.Show
    Set obj = Nothing
End Sub

'/**
' * @brief 権限チェック（管理者/品管/ライン長）
' * @param i_KBN 権限区分
' * @return String 権限がある場合は氏名、なければ空文字
' * @note DBから該当権限の氏名を取得
' */
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
