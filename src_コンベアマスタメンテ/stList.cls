VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'/**
' * @file stList.cls
' * @brief 工程リストシートモジュール (UI 層)
' * @note ユーザー操作の受付・入力チェック・保存・終了処理などを担当
' */

Option Explicit

'/**
' * @brief 選択ボタン押下時の処理
' */
Sub stList_Select_Click()
    On Error GoTo ErrorHandler
    '再表示処理
    If fncChkChange Then
        If MsgBox("変更内容が保存されていませんが、よろしいですか", vbOKCancel, Bas_Configuration.SYSTEM_NAME) = vbCancel Then Exit Sub
    End If
    Call subMain
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.stList_Select_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("選択処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 保存ボタン押下時の処理
' */
Sub stList_Update_Click()
    On Error GoTo ErrorHandler
    If Not fncInputCheck Then Exit Sub
    Call subBeforeEdit
    Call Bas_DaoConveyor.UpdateConveyorList(Me, stWork, RW_FR)
    Call subMain
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.stList_Update_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("保存処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 終了ボタン押下時の処理
' */
Sub stList_End_Click()
    On Error GoTo ErrorHandler
    If fncChkChange Then
        If MsgBox("変更内容が保存されていませんが、よろしいですか", vbOKCancel, Bas_Configuration.SYSTEM_NAME) = vbCancel Then Exit Sub
    End If
    Me.Cells(1, 1).Copy
    Application.CutCopyMode = False
    If Workbooks.Count = 1 Then Application.Quit
    Application.ThisWorkbook.Close False
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.stList_End_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("終了処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief セル変更時の処理
' */
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    If Target.Row < RW_FR Then Exit Sub
    If Target.Column > 3 Then Exit Sub '3列目(名称)まで
    If Not Me.Cells(Target.Row, Target.Column) = "" Then
        '枠線を引く(A～C列)
        Me.Range(Me.Cells(RW_FR, 1), Me.Cells(Target.Row, 3)).Borders.LineStyle = xlContinuous
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.Worksheet_Change: " & Err.Number & " - " & Err.Description
    Call MsgBox("セル変更処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief セル選択変更時の処理
' */
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrorHandler
    If Target.Count > 1 Then Exit Sub
    If Target.Row < RW_FR Then Exit Sub
    'A列(工程)選択時にフォームを表示
    If Target.Column = 1 Then
        Call subOpen
        If Not P_Regist Then GoTo Exit_
        Me.Cells(Target.Row, 1) = P_SelItem
    End If
Exit_:
    'Me.Cells(RW_FR - 1, 1).Select '強制移動させたい場合はコメントアウト解除
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.Worksheet_SelectionChange: " & Err.Number & " - " & Err.Description
    Call MsgBox("セル選択処理中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

'/**
' * @brief 選択フォームの表示
' */
Private Sub subOpen()
    On Error GoTo ErrorHandler
    Dim obj As New frmSelectList
    obj.Show
    Set obj = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] stList.subOpen: " & Err.Number & " - " & Err.Description
    Call MsgBox("選択フォーム表示中にエラーが発生しました。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub

Private Function fncInputCheck() As Boolean
    Dim strWk       As String
    Dim strMSG      As String
    Dim lRow        As Long
    Dim lRow2       As Long
    Dim MaxRow      As Long
    
    MaxRow = fncGetMaxRow
    
    '色クリア
    Me.Range(Me.Cells(RW_FR, 1), Me.Cells(MaxRow, 3)).Interior.Color = xlNone
    
    For lRow = RW_FR To MaxRow
        If Trim(Me.Cells(lRow, 1)) <> "" Or Trim(Me.Cells(lRow, 2)) <> "" Or Trim(Me.Cells(lRow, 3)) <> "" Then
            '1.工程チェック
            strWk = fncGetKTCD(Trim(Me.Cells(lRow, 1)))
            If strWk = "" Then strMSG = "工程が正しく入力されていません": Me.Range(Me.Cells(lRow, 1), Me.Cells(lRow, 3)).Interior.Color = RGB(255, 153, 204): Exit For
            
            '2.Noチェック
            strWk = StrConv(Trim(Me.Cells(lRow, 2)), vbNarrow)
            If Not IsNumeric(strWk) Then strMSG = "Noには1以上の数値を入力してください": Me.Range(Me.Cells(lRow, 1), Me.Cells(lRow, 3)).Interior.Color = RGB(255, 153, 204): Exit For
            If Val(strWk) = 0 Then strMSG = "Noには1以上の数値を入力してください": Me.Range(Me.Cells(lRow, 1), Me.Cells(lRow, 3)).Interior.Color = RGB(255, 153, 204): Exit For
            
            '3.名称チェック
            strWk = Trim(Me.Cells(lRow, 3))
            If strWk = "" Then strMSG = "名称を入力してください": Me.Range(Me.Cells(lRow, 1), Me.Cells(lRow, 3)).Interior.Color = RGB(255, 153, 204): Exit For
        End If
    Next
    
    If strMSG <> "" Then Call MsgBox(strMSG, vbOKOnly): Exit Function
    
    '重複チェック (工程 + No)
    For lRow = RW_FR To MaxRow
        For lRow2 = RW_FR To MaxRow
            If Not lRow = lRow2 Then
                If Trim(Me.Cells(lRow, 1)) <> "" And Trim(Me.Cells(lRow2, 1)) <> "" Then
                    If fncGetKTCD(Trim(Me.Cells(lRow, 1))) = fncGetKTCD(Trim(Me.Cells(lRow2, 1))) And _
                       Val(StrConv(Trim(Me.Cells(lRow, 2)), vbNarrow)) = Val(StrConv(Trim(Me.Cells(lRow2, 2)), vbNarrow)) Then
                        
                        Me.Range(Me.Cells(lRow, 1), Me.Cells(lRow, 3)).Interior.Color = RGB(255, 153, 204)
                        Me.Range(Me.Cells(lRow2, 1), Me.Cells(lRow2, 3)).Interior.Color = RGB(255, 153, 204)
                        strMSG = "工程とNoの組み合わせが重複しています": Exit For
                    End If
                End If
            End If
        Next
        If strMSG <> "" Then Exit For
    Next
    
    If strMSG <> "" Then Call MsgBox(strMSG, vbOKOnly): Exit Function

    fncInputCheck = True
End Function

Private Function fncChkChange() As Boolean
    Dim lRow        As Long
    Dim lRow2       As Long
    Dim MaxRow      As Long
    Dim bEdit       As Boolean
    
    fncChkChange = False
    'データが無い場合はチェック不要
    If stWork.Cells(RW_FR, 1) = "" And Me.Cells(RW_FR, 1) = "" Then Exit Function
    
    MaxRow = fncGetMaxRow
    
    '【変更・削除チェック】
    lRow = RW_FR
    Do While Not stWork.Cells(lRow, 1) = ""
        bEdit = True
        For lRow2 = RW_FR To MaxRow
            '比較：工程(1)、No(2)、名称(3)
            If stWork.Cells(lRow, 1) = Me.Cells(lRow2, 1) And _
               Val(stWork.Cells(lRow, 2)) = Val(Me.Cells(lRow2, 2)) And _
               stWork.Cells(lRow, 3) = Me.Cells(lRow2, 3) Then
                bEdit = False
                Exit For
            End If
        Next
        If bEdit Then
            fncChkChange = True
            Exit Function
        End If
        lRow = lRow + 1
    Loop
    
    '【追加チェック】
    For lRow2 = RW_FR To MaxRow
        If Me.Cells(lRow2, 1) <> "" Then
            bEdit = True
            lRow = RW_FR
            Do While Not stWork.Cells(lRow, 1) = ""
                '工程(1)とNo(2)が一致する行があれば「既存」とみなす
                If Me.Cells(lRow2, 1) = stWork.Cells(lRow, 1) And _
                   Val(Me.Cells(lRow2, 2)) = Val(stWork.Cells(lRow, 2)) Then
                    bEdit = False
                    Exit Do
                End If
                lRow = lRow + 1
            Loop
            If bEdit Then
                fncChkChange = True
                Exit Function
            End If
        End If
    Next
End Function
