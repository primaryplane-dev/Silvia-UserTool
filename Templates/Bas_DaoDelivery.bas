Attribute VB_Name = "Bas_DaoDelivery"
'/**
' * @file Bas_DaoDelivery.bas
' * @brief 【SAMPLE】配送データアクセス層 (DAL)
' * @note 配送情報の取得および更新を行うためのサンプル実装
' */

Option Explicit

'/**
' * @brief 指定された ID の配送データを取得する
' * @param lngDeliveryID 取得対象の配送 ID
' * @return Result オブジェクト (Data プロパティに Delivery オブジェクトを格納)
' */
Public Function GetDeliveryById(ByVal lngDeliveryID As Long) As Result
    On Error GoTo ErrorHandler

    Dim objResult As Result
    Set objResult = New Result
    
    ' DB 接続の取得
    Dim objConnection As Object
    Set objConnection = Bas_DbConnection.GetConnection()
    
    If objConnection Is Nothing Then
        objResult.IsSuccess = False
        objResult.Message = "データベースに接続できません。"
        Set GetDeliveryById = objResult
        Exit Function
    End If

    ' SQL 構築
    Dim strSql As String
    strSql = "SELECT DeliveryID, DestinationName, DeliveryDate, Status " & _
             "FROM Deliveries " & _
             "WHERE DeliveryID = " & lngDeliveryID

    ' Recordset 実行
    Dim objRecordset As Object
    Set objRecordset = CreateObject("ADODB.Recordset")
    Call objRecordset.Open(strSql, objConnection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY)

    ' データのマッピング
    If Not objRecordset.EOF Then
        Dim objDelivery As Delivery
        Set objDelivery = New Delivery
        
        objDelivery.DeliveryID = objRecordset.Fields("DeliveryID").Value
        objDelivery.DestinationName = objRecordset.Fields("DestinationName").Value & ""
        
        If Not IsNull(objRecordset.Fields("DeliveryDate").Value) Then
            objDelivery.DeliveryDate = CDate(objRecordset.Fields("DeliveryDate").Value)
        End If
        
        objDelivery.Status = objRecordset.Fields("Status").Value & ""

        objResult.IsSuccess = True
        Set objResult.Data = objDelivery
    Else
        objResult.IsSuccess = False
        objResult.Message = "対象の配送データが見つかりませんでした。"
    End If

    ' リソース解放
    Call objRecordset.Close
    Set objRecordset = Nothing

    Set GetDeliveryById = objResult
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_DaoDelivery.GetDeliveryById: " & Err.Number & " - " & Err.Description
    
    objResult.IsSuccess = False
    objResult.Message = "データ取得処理で予期せぬエラーが発生しました。"
    
    If Not objRecordset Is Nothing Then
        If objRecordset.State = AD_STATE_OPEN Then
            Call objRecordset.Close
        End If

        Set objRecordset = Nothing
    End If
    
    Set GetDeliveryById = objResult
End Function

'/**
' * @brief 指定された ID の配送ステータスを更新する (トランザクション対応)
' * @param lngDeliveryID 更新対象の ID
' * @param strNewStatus 新しいステータス名
' * @return Result 処理結果
' */
Public Function UpdateStatus(ByVal lngDeliveryID As Long, ByVal strNewStatus As String) As Result
    On Error GoTo ErrorHandler

    Dim objResult As Result
    Set objResult = New Result
    
    Dim objConnection As Object
    Set objConnection = Bas_DbConnection.GetConnection()
    
    If objConnection Is Nothing Then
        objResult.IsSuccess = False
        objResult.Message = "データベースに接続できないため、更新できませんでした。"
        Set UpdateStatus = objResult
        Exit Function
    End If

    ' トランザクション開始
    Call objConnection.BeginTrans

    ' Command オブジェクトの作成 (パラメータクエリの実装)
    Dim objCommand As Object
    Set objCommand = CreateObject("ADODB.Command")
    
    With objCommand
        Set .ActiveConnection = objConnection
        .CommandText = "UPDATE Deliveries SET Status = ? WHERE DeliveryID = ?"
        .CommandType = Bas_DataConstants.AD_CMD_TEXT
        
        ' パラメータの追加 (順序は SQL 文の ? の順に合わせる)
        ' 1. Status (文字列型, サイズは DB の定義に合わせて指定。ここでは仮に 50 とする)
        Call .Parameters.Append(.CreateParameter("Status", Bas_DataConstants.AD_VAR_CHAR, Bas_DataConstants.AD_PARAM_INPUT, 50, strNewStatus))
        
        ' 2. DeliveryID (整数型)
        Call .Parameters.Append(.CreateParameter("DeliveryID", Bas_DataConstants.AD_INTEGER, Bas_DataConstants.AD_PARAM_INPUT, , lngDeliveryID))
        
        ' SQL 実行
        Call .Execute
    End With
    
    Set objCommand = Nothing

    ' コミット (確定)
    Call objConnection.CommitTrans

    objResult.IsSuccess = True
    objResult.Message = "ステータスを更新しました。"
    Set UpdateStatus = objResult
    Exit Function

ErrorHandler:
    ' エラー時はロールバック (取り消し)
    If Not objConnection Is Nothing Then
        If objConnection.State = AD_STATE_OPEN Then
            Call objConnection.RollbackTrans
        End If
    End If

    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_DaoDelivery.UpdateStatus: " & Err.Number & " - " & Err.Description
    
    objResult.IsSuccess = False
    objResult.Message = "データの更新中にエラーが発生したため、変更は破棄されました。"
    Set UpdateStatus = objResult
End Function
