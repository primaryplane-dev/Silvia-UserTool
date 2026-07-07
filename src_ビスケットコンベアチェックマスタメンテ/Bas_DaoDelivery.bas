'' コンベア一覧取得（SBFP01テーブル）
Public Function GetConveyorList() As Object
    On Error GoTo ErrorHandler
    Dim objResult As Object
    Set objResult = CreateObject("ADODB.Recordset")
    Dim CN As Object: Set CN = Bas_DbConnection.GetConnection()
    Dim strSQL As String
    strSQL = "SELECT BFKTCD, BFCVNO, BFCVNM FROM LIBSMF17.SBFP01 WHERE BFDELT='' ORDER BY BFKTCD, BFCVNO"
    objResult.Open strSQL, CN, 3, 1 ' adOpenStatic=3, adLockReadOnly=1
    Set GetConveyorList = objResult
    Exit Function
ErrorHandler:
    Set GetConveyorList = Nothing
End Function
Attribute VB_Name = "Bas_DaoDelivery"
'/**
' * @file Bas_DaoDelivery.bas
' * @brief 配送データアクセス層 (DAL)
' * @note 配送情報の取得および更新を行うための実装
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
        objResult.Message = "データベースに接続できません。"
        Set UpdateStatus = objResult
        Exit Function
    End If

    objConnection.BeginTrans
    Dim strSql As String
    strSql = "UPDATE Deliveries SET Status = '" & strNewStatus & "' WHERE DeliveryID = " & lngDeliveryID
    On Error Resume Next
    objConnection.Execute strSql
    If Err.Number = 0 Then
        objConnection.CommitTrans
        objResult.IsSuccess = True
        objResult.Message = "ステータスを更新しました。"
    Else
        objConnection.RollbackTrans
        objResult.IsSuccess = False
        objResult.Message = "ステータス更新に失敗しました: " & Err.Description
    End If
    On Error GoTo 0
    Set UpdateStatus = objResult
    Exit Function

ErrorHandler:
    objConnection.RollbackTrans
    objResult.IsSuccess = False
    objResult.Message = "更新処理で予期せぬエラーが発生しました。"
    Set UpdateStatus = objResult
End Function
