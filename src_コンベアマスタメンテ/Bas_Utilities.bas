'' AS/400の8桁日付（yyyyMMdd）をDate型に変換
Public Function ConvertDate(ByVal strDate As String) As Date
    On Error GoTo ErrorHandler
    strDate = Trim(strDate)
    If Len(strDate) = 8 And IsNumeric(strDate) Then
        Dim strFormatted As String
        strFormatted = Left(strDate, 4) & "/" & Mid(strDate, 5, 2) & "/" & Right(strDate, 2)
        If IsDate(strFormatted) Then
            ConvertDate = CDate(strFormatted)
            Exit Function
        End If
    End If
    ConvertDate = CDate("1900/01/01")
    Exit Function
ErrorHandler:
    ConvertDate = CDate("1900/01/01")
End Function
Attribute VB_Name = "Bas_Utilities"
Option Explicit

' 設定ファイル（Key=Value形式）から値を取得
Public Function GetIniValue(ByVal strKey As String, ByVal strDefault As String, ByVal strFilePath As String) As String
    On Error GoTo ErrorHandler
    GetIniValue = strDefault
    Dim objFileSystemObject As Object
    Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If objFileSystemObject.FileExists(strFilePath) = False Then GoTo Finally
    Dim objStream As Object
    Set objStream = objFileSystemObject.OpenTextFile(strFilePath, 1, False, -2) ' -2: システム既定(Shift-JIS)
    Dim strLine As String
    Do While Not objStream.AtEndOfStream
        strLine = Trim(objStream.ReadLine)
        If strLine <> "" And InStr(strLine, "=") > 0 Then
            If Split(strLine, "=")(0) = strKey Then
                GetIniValue = Split(strLine, "=")(1)
                Exit Do
            End If
        End If
    Loop
    objStream.Close
Finally:
    Set objStream = Nothing
    Set objFileSystemObject = Nothing
    Exit Function
ErrorHandler:
    GetIniValue = strDefault
    Resume Finally
End Function
