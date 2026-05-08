VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenu 
   Caption         =   "条件指定"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13290
   OleObjectBlob   =   "frmMenu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' フォーム初期化
' ==============================================================================
Private Sub UserForm_Initialize()
    ' 日付の初期値をセット
    Me.txtDate.Value = Format(Date, "yyyy/mm/dd")
    
    ' 区分（デフォルトで「成型」を選んだ状態にする）
    Me.optSeikei.Value = True
    
    ' 商品名コンボボックスの設定（3列：商品名, 日番号, ラインNo）
    With Me.cmbHinm
        .ColumnCount = 3
        .ColumnWidths = "100pt;0pt;0pt" ' 2列目以降は隠す
        .Style = fmStyleDropDownList    ' 手入力禁止
    End With
    
    ' 初期日付でリストを作成
    Call subMakeComboHINM
End Sub

' ==============================================================================
' 実行ボタン
' ==============================================================================
Private Sub cmdRun_Click()
    ' 1. 入力チェック
    If Not IsDate(Me.txtDate.Value) Then
        MsgBox "正しい日付を入力してください。", vbExclamation
        Me.txtDate.SetFocus
        Exit Sub
    End If
    
    If Me.cmbHinm.Value = "" Or Me.cmbHinm.Value = "該当データなし" Or Me.cmbHinm.Value = "SQLエラー" Then
        MsgBox "有効な商品名を選択してください。" & vbCrLf & "（データが存在しない日付では実行できません）", vbExclamation
        Me.cmbHinm.SetFocus
        Exit Sub
    End If
    

    ' 2. 標準モジュールのパブリック変数に値をセット
    P_DATE = CDate(Me.txtDate.Value)
    
    ' 商品名、日番号、ラインNoを取得
    P_HINM = Me.cmbHinm.List(Me.cmbHinm.ListIndex, 0) & ""
    P_HINO = Me.cmbHinm.List(Me.cmbHinm.ListIndex, 1) & ""
    P_KJNO = Me.cmbHinm.List(Me.cmbHinm.ListIndex, 2) & ""
    
    ' 区分判定
    If Me.optReikyaku.Value = True Then
        P_ChkKBN = 2 ' 冷却
    Else
        P_ChkKBN = 1 ' 成型
    End If
    
    ' 3. メイン処理の呼び出し
    Me.Hide
    DoEvents
    
    Call subMain  ' 標準モジュールのメイン処理を実行
    
    ' 4. 処理終了後にフォームを閉じる
    Unload Me
End Sub

' キャンセルボタン
Private Sub cmdCancel_Click()
    Unload Me
End Sub

' ==============================================================================
' カレンダー・日付関連
' ==============================================================================
Private Sub cmdCalD_Click()
    Call subOpenCalendar
    
    ' カレンダーで日付が選択された場合のみ更新
    If P_CalendarSelected Then
        Me.txtDate.Value = Format(P_calDATE, "yyyy/mm/dd")
        ' 日付が変わったので商品リストを再取得
        Call subMakeComboHINM
    End If
End Sub

Private Sub subOpenCalendar()
    P_CalendarSelected = False
    Dim obj As New frmCalendar
    obj.Show
    Set obj = Nothing
End Sub

' 日付を手入力してEnter等で抜けた時にリスト更新
Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(Me.txtDate.Value) Then
        Call subMakeComboHINM
    End If
End Sub

' ==============================================================================
' DBからその日の商品リストを取得する処理
' ==============================================================================
Private Sub subMakeComboHINM()
    Dim CN As Object
    ' コンボボックスをクリア
    Me.cmbHinm.Clear
    ' 日付が不正なら終了
    If Not IsDate(Me.txtDate.Value) Then Exit Sub
    Dim yyyymmdd As String
    yyyymmdd = Format(Me.txtDate.Value, "yyyymmdd")
    Dim RS As Object
    Set RS = Bas_LogicConveyor.GetHinmListWithLogic(yyyymmdd)
    If RS Is Nothing Then
        Me.cmbHinm.AddItem "SQLエラー"
        Exit Sub
    End If
    If Not RS.EOF Then
        Do While Not RS.EOF
            Dim sName As String
            sName = RS("BGHINM") & ""
            If sName = "" Then sName = "(名称未設定)"
            Me.cmbHinm.AddItem sName
            Me.cmbHinm.List(Me.cmbHinm.ListCount - 1, 1) = RS("BGHINO") & ""
            Me.cmbHinm.List(Me.cmbHinm.ListCount - 1, 2) = RS("BGKJNO") & ""
            RS.MoveNext
        Loop
        Me.cmbHinm.ListIndex = 0
    Else
        Me.cmbHinm.AddItem "該当データなし"
    End If
    RS.Close
    Set RS = Nothing
    Exit Sub

ErrConnect:
    MsgBox "DB接続に失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub
