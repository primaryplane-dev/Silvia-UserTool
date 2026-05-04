Attribute VB_Name = "Bas_Configuration"
Option Explicit

' システム名
Public Const SYSTEM_NAME As String = "コンベア一覧表システム"

' DB接続タイムアウト（秒）
Public Const DB_CONNECTION_TIMEOUT As Long = 30
Public Const DB_COMMAND_TIMEOUT As Long = 60

' ログ日付フォーマット
Public Const LOG_DATE_FORMAT As String = "yyyy/MM/dd HH:mm:ss"

' 設定ファイルの文字コード
Public Const INI_FILE_CHARSET As String = "Shift_JIS"

' DB接続文字列を取得
Public Function GetConnectionString() As String
    On Error GoTo ErrorHandler
    Dim strIniFilePath As String
    strIniFilePath = ThisWorkbook.Path & "\設定.ini"
    Dim strDataSource As String
    Dim strUserID As String
    Dim strPassword As String
    strDataSource = Bas_Utilities.GetIniValue("DSN", "", strIniFilePath)
    strUserID = Bas_Utilities.GetIniValue("UID", "", strIniFilePath)
    strPassword = Bas_Utilities.GetIniValue("PWD", "", strIniFilePath)
    If strDataSource = "" Or strUserID = "" Then
        Call MsgBox("設定ファイル(設定.ini)にDB接続情報が正しく設定されていません。", vbCritical, SYSTEM_NAME)
        GetConnectionString = ""
        Exit Function
    End If
    GetConnectionString = "Provider=IBMDA400;Data Source=" & strDataSource & ";User ID=" & strUserID & ";Password=" & strPassword & ";"
    Exit Function
ErrorHandler:
    Call MsgBox("システムの接続設定の読み込みに失敗しました。管理者へ連絡してください。", vbCritical, SYSTEM_NAME)
    GetConnectionString = ""
End Function
