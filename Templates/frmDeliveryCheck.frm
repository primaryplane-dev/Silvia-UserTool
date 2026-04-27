VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDeliveryCheck 
   Caption         =   "画面サンプル"
   ClientHeight    =   2076
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3504
   OleObjectBlob   =   "frmDeliveryCheck.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDeliveryCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**
' * @file frmDeliveryCheck.frm
' * @brief 【SAMPLE】配送状況確認画面 (UI 層)
' * @note ユーザー入力の受け付けと、ロジック層の呼び出し、結果の表示のみを行う
' */

Option Explicit

'/**
' * @brief 確認ボタン押下時の処理
' */
Private Sub cmdCheck_Click()
    On Error GoTo ErrorHandler

    ' 入力値の簡易バリデーション (UI 層の責任)
    If Not IsNumeric(Me.txtDeliveryID.Text) Then
        Call MsgBox("配送 ID には数値を入力してください。", vbExclamation, Bas_Configuration.SYSTEM_NAME)
        Exit Sub
    End If

    Dim lngDeliveryID As Long
    lngDeliveryID = CLng(Me.txtDeliveryID.Text)

    ' ビジネスロジック層の呼び出し
    Dim objResult As Result
    Set objResult = Bas_LogicDelivery.CheckDeliveryStatus(lngDeliveryID)

    ' 処理結果の判定と画面への反映
    If objResult.IsSuccess Then
        ' 成功時は Data プロパティから Delivery オブジェクトを取り出して表示
        Dim objDelivery As Delivery
        Set objDelivery = objResult.Data
        
        Me.lblDestinationName.Caption = objDelivery.DestinationName
        Me.lblStatus.Caption = objDelivery.Status
        
        Call MsgBox(objResult.Message, vbInformation, Bas_Configuration.SYSTEM_NAME)
    Else
        ' 失敗時はメッセージを表示し、画面の項目をクリア
        Me.lblDestinationName.Caption = ""
        Me.lblStatus.Caption = ""
        
        Call MsgBox(objResult.Message, vbExclamation, Bas_Configuration.SYSTEM_NAME)
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Frm_DeliveryCheck.cmdCheck_Click: " & Err.Number & " - " & Err.Description
    Call MsgBox("処理の実行中にエラーが発生しました。再度実行しても改善しない場合は、システム管理者へ連絡してください。", vbCritical, Bas_Configuration.SYSTEM_NAME)
End Sub
