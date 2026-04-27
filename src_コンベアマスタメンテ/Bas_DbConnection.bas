Attribute VB_Name = "Bas_DbConnection"
Option Explicit

Private M_objConnection As Object

' DB接続を取得または新規作成
Public Function GetConnection() As Object
    On Error GoTo ErrorHandler
    If Not M_objConnection Is Nothing Then
        If M_objConnection.State = 1 Then ' AD_STATE_OPEN = 1
            Set GetConnection = M_objConnection
            Exit Function
        End If
    End If
    Set M_objConnection = CreateObject("ADODB.Connection")
    With M_objConnection
        .ConnectionString = Bas_Configuration.GetConnectionString()
        .ConnectionTimeout = Bas_Configuration.DB_CONNECTION_TIMEOUT
        .CommandTimeout = Bas_Configuration.DB_COMMAND_TIMEOUT
        .CursorLocation = 3 ' AD_USE_CLIENT = 3
        .Open
    End With
    Set GetConnection = M_objConnection
    Exit Function
ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_DbConnection.GetConnection: " & Err.Number & " - " & Err.Description
    Call MsgBox("DB接続に失敗しました。ネットワークや設定を確認してください。", vbCritical, Bas_Configuration.SYSTEM_NAME)
    Set GetConnection = Nothing
End Function

' DB接続を安全に閉じる
Public Sub CloseConnection()
    On Error GoTo ErrorHandler
    If Not M_objConnection Is Nothing Then
        If M_objConnection.State = 1 Then
            M_objConnection.Close
        End If
        Set M_objConnection = Nothing
    End If
    Exit Sub
ErrorHandler:
    Set M_objConnection = Nothing
End Sub
