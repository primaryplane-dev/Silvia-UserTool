Attribute VB_Name = "Bas_Configuration"
'/**
' * @file Bas_Configuration.bas
' * @brief 環境設定および共通定数管理モジュール
' */

Option Explicit

'/** @brief システム名称 (メッセージボックスのタイトル等で使用) */
Public Const SYSTEM_NAME As String = "サンプルシステム"

'/** @brief データベース接続タイムアウト (秒) */
Public Const DB_CONNECTION_TIMEOUT As Long = 30
'/** @brief データベースコマンド実行タイムアウト (秒) */
Public Const DB_COMMAND_TIMEOUT As Long = 60

'/** @brief ログ出力用などの標準日時フォーマット */
Public Const LOG_DATE_FORMAT As String = "yyyy/MM/dd HH:mm:ss"

'/** @brief 設定ファイルの文字コード */
Public Const INI_FILE_CHARSET As String = "Shift_JIS"

'/**
' * @brief データベース接続文字列を取得する
' * @return String 接続文字列
' */
Public Function GetConnectionString() As String
    On Error GoTo ErrorHandler
    Dim strIniFilePath As String
    ' このブックと同じ階層にある 設定.ini を指定
    strIniFilePath = ThisWorkbook.Path & "\設定.ini"
    Dim strDataSource As String
    Dim strUserID As String
    Dim strPassword As String
    ' セクションなしの独自パーサーを使用して設定情報を読み込む
    strDataSource = Bas_Utilities.GetIniValue("DSN", "", strIniFilePath)
    strUserID = Bas_Utilities.GetIniValue("UID", "", strIniFilePath)
    strPassword = Bas_Utilities.GetIniValue("PWD", "", strIniFilePath)
    ' 必須設定のチェック
    If strDataSource = "" Or strUserID = "" Then
        Call MsgBox("設定ファイル (設定.ini) にデータベース接続情報が正しく設定されていません。", vbCritical, SYSTEM_NAME)
        GetConnectionString = ""
        Exit Function
    End If
    ' IBMDA400 (AS/400) 向けの接続文字列を構築
    GetConnectionString = "Provider=IBMDA400;Data Source=" & strDataSource & ";User ID=" & strUserID & ";Password=" & strPassword & ";"
    Exit Function
ErrorHandler:
    Call MsgBox("システムの接続設定の読み込みに失敗しました。システム管理者へ連絡してください。", vbCritical, SYSTEM_NAME)
    GetConnectionString = ""
End Function
