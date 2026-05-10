Attribute VB_Name = "Bas_Utilities"
'/**
' * @file Bas_Utilities.bas
' * @brief 汎用ユーティリティ関数群
' */

Option Explicit

'/**
' * @brief INIファイルから値を取得する
' * @param key 取得するキー名
' * @param defaultValue デフォルト値
' * @param iniFilePath INIファイルのパス
' * @return String キーに対応する値（見つからない場合はデフォルト値）
' */
Public Function GetIniValue(ByVal key As String, ByVal defaultValue As String, ByVal iniFilePath As String) As String
    On Error GoTo ErrorHandler
    Dim fileNo As Integer: fileNo = FreeFile
    Dim line As String, pos As Long
    GetIniValue = defaultValue
    Open iniFilePath For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, line
        If InStr(line, "=") > 0 Then
            If LCase(Trim(Split(line, "=")(0))) = LCase(key) Then
                GetIniValue = Trim(Mid(line, InStr(line, "=") + 1))
                Exit Do
            End If
        End If
    Loop
    Close #fileNo
    Exit Function
ErrorHandler:
    GetIniValue = defaultValue
    If fileNo > 0 Then Close #fileNo
End Function

'/**
' * @brief 日付文字列のフォーマット変換
' * @param dt 日付型または文字列
' * @param format フォーマット文字列
' * @return String フォーマット済み日付
' */
Public Function FormatDateString(ByVal dt As Variant, ByVal format As String) As String
    On Error Resume Next
    FormatDateString = Format(dt, format)
End Function

'/**
' * @brief エラーログを標準出力に記録
' * @param msg ログメッセージ
' */
Public Sub LogError(ByVal msg As String)
    Debug.Print "[ERROR] " & Format(Now, "yyyy/MM/dd HH:mm:ss") & ": " & msg
End Sub
