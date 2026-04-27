Attribute VB_Name = "Bas_LogicDelivery"
'/**
' * @file Bas_LogicDelivery.bas
' * @brief 【SAMPLE】配送業務ロジック層 (BLL)
' * @note 配送に関する業務ルールを適用するサンプル実装
' */

Option Explicit

'/**
' * @brief 配送データのステータスをチェックする
' * @param lngDeliveryID 対象の配送 ID
' * @return Result オブジェクト
' */
Public Function CheckDeliveryStatus(ByVal lngDeliveryID As Long) As Result
    On Error GoTo ErrorHandler

    Dim objResult As Result
    Set objResult = New Result

    ' DAO を呼び出してデータを取得
    Dim objDaoResult As Result
    Set objDaoResult = Bas_DaoDelivery.GetDeliveryById(lngDeliveryID)

    ' データ取得に失敗した場合はそのまま結果を返す
    If Not objDaoResult.IsSuccess Then
        Set CheckDeliveryStatus = objDaoResult
        Exit Function
    End If

    ' 業務ルールの適用 (ステータスチェック)
    Dim objDelivery As Delivery
    Set objDelivery = objDaoResult.Data

    If objDelivery.Status = "キャンセル" Then
        objResult.IsSuccess = False
        objResult.Message = "この配送は既にキャンセルされています。"
        Set CheckDeliveryStatus = objResult
        Exit Function
    End If

    ' 正常終了
    objResult.IsSuccess = True
    objResult.Message = "正常な配送データです。"
    Set objResult.Data = objDelivery
    
    Set CheckDeliveryStatus = objResult
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_LogicDelivery.CheckDeliveryStatus: " & Err.Number & " - " & Err.Description
    
    objResult.IsSuccess = False
    objResult.Message = "業務ロジックの実行中にエラーが発生しました。"
    Set CheckDeliveryStatus = objResult
End Function
