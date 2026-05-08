Attribute VB_Name = "Bas_DbConnection"
Option Explicit

'/**
' * @file Bas_DbConnection.bas
' * @brief データベース接続・トランザクション管理モジュール
' */

'/** @brief 現在のデータベース接続オブジェクトを保持するモジュール変数 */
Private M_objConnection As Object

'/**
' * @brief データベース接続を取得または新規作成する
' * @return Object ADODB.Connection オブジェクト (失敗時は Nothing)
' */
Public Function GetConnection() As Object
    On Error GoTo ErrorHandler

    ' 既に接続が開かれている場合はそれを返す
    If Not M_objConnection Is Nothing Then
        If M_objConnection.State = AD_STATE_OPEN Then
            Set GetConnection = M_objConnection
            Exit Function
        End If
    End If

    ' 新規接続の作成
    Set M_objConnection = CreateObject("ADODB.Connection")
    With M_objConnection
        .ConnectionString = Bas_Configuration.GetConnectionString()
        .ConnectionTimeout = Bas_Configuration.DB_CONNECTION_TIMEOUT
        .CommandTimeout = Bas_Configuration.DB_COMMAND_TIMEOUT
        .CursorLocation = AD_USE_CLIENT
        Call .Open
    End With

    Set GetConnection = M_objConnection
    Exit Function

ErrorHandler:
    ' タイムスタンプ付きでイミディエイトウィンドウにログ出力
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_DbConnection.GetConnection: " & Err.Number & " - " & Err.Description
    
    Call MsgBox("データベースへの接続に失敗しました。ネットワーク状態や設定を確認してください。", vbCritical, Bas_Configuration.SYSTEM_NAME)
    
    Set GetConnection = Nothing
End Function

'/**
' * @brief データベース接続を安全に閉じて解放する
' */
Public Sub CloseConnection()
    On Error GoTo ErrorHandler
    
    If Not M_objConnection Is Nothing Then
        If M_objConnection.State = AD_STATE_OPEN Then
            Call M_objConnection.Close
        End If
        Set M_objConnection = Nothing
    End If
    
    Exit Sub

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_DbConnection.CloseConnection: " & Err.Number & " - " & Err.Description
    Set M_objConnection = Nothing
End Sub
