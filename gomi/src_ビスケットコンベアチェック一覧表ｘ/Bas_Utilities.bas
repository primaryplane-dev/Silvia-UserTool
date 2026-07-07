Option Explicit

'/**
' * @brief AS/400 の 8桁数値日付 (yyyyMMdd) を VBA の Date 型に変換する
' * @param strDate 8桁の文字列または数値
' * @return Date 変換後の日付。変換失敗時は 1900/01/01 を返す。
' */
Public Function ConvertDate(ByVal strDate As String) As Date
    On Error GoTo ErrorHandler

    ' 固定長文字列の末尾空白を除去 (規約 5.3 対応)
    strDate = Trim(strDate)

    ' 8桁の数値であるかチェック
    If Len(strDate) = 8 And IsNumeric(strDate) Then
        Dim strFormatted As String
        strFormatted = Left(strDate, 4) & "/" & Mid(strDate, 5, 2) & "/" & Right(strDate, 2)

        If IsDate(strFormatted) Then
            ConvertDate = CDate(strFormatted)
            Exit Function
        End If
    End If

    ' 不正な値の場合はデフォルト日付を返す
    ConvertDate = CDate("1900/01/01")
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_Utilities.ConvertDate: " & Err.Number & " - " & Err.Description
    ConvertDate = CDate("1900/01/01")
End Function


'/**
' * @brief 設定ファイル (セクションなし Key=Value 形式) から値を取得する
' * @param strKey 取得したい項目のキー名
' * @param strDefault 取得失敗時や未設定時のデフォルト値
' * @param strFilePath 設定ファイルのフルパス
' * @return String 取得した値
' */
Public Function GetIniValue(ByVal strKey As String, ByVal strDefault As String, ByVal strFilePath As String) As String
    On Error GoTo ErrorHandler

    GetIniValue = strDefault

    Dim objFileSystemObject As Object
    Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If objFileSystemObject.FileExists(strFilePath) = False Then
        GoTo Finally
    End If

    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")

    objStream.Type = Bas_DataConstants.AD_TYPE_TEXT
    objStream.Charset = Bas_Configuration.INI_FILE_CHARSET

    Call objStream.Open
    Call objStream.LoadFromFile(strFilePath)

    Dim strFileText As String
    strFileText = objStream.ReadText

    strFileText = Replace(strFileText, vbCrLf, vbLf)
    strFileText = Replace(strFileText, vbCr, vbLf)

    Dim arrStrLines() As String
    arrStrLines = Split(strFileText, vbLf)

    Dim lngLineIndex As Long
    For lngLineIndex = LBound(arrStrLines) To UBound(arrStrLines)
        Dim strLine As String
        strLine = Trim(arrStrLines(lngLineIndex))

        If IsIniValueLine(strLine) Then
            Dim arrStrParts() As String
            arrStrParts = Split(strLine, "=", 2)
            If StrComp(Trim(arrStrParts(0)), Trim(strKey), vbTextCompare) = 0 Then
                GetIniValue = Trim(arrStrParts(1))
                Exit For
            End If
        End If
    Next lngLineIndex

Finally:
    If Not objStream Is Nothing Then
        If objStream.State = Bas_DataConstants.AD_STATE_OPEN Then
            Call objStream.Close
        End If
        Set objStream = Nothing
    End If
    If Not objFileSystemObject Is Nothing Then
        Set objFileSystemObject = Nothing
    End If
    
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_Utilities.GetIniValue: " & Err.Number & " - " & Err.Description
    Resume Finally
End Function

'/**
' * @brief INI 設定値として解析対象にする行かどうかを判定する
' * @param strLine [in] 判定対象の 1 行
' * @return Boolean 解析対象の場合 True、それ以外の場合 False
' */
Private Function IsIniValueLine(ByVal strLine As String) As Boolean
    On Error GoTo ErrorHandler

    IsIniValueLine = False

    If Len(strLine) = 0 Then
        Exit Function
    End If

    If Left(strLine, 1) = "#" Then
        Exit Function
    End If

    If Left(strLine, 1) = ";" Then
        Exit Function
    End If

    If InStr(strLine, "=") = 0 Then
        Exit Function
    End If

    IsIniValueLine = True

    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_Utilities.IsIniValueLine: " & Err.Number & " - " & Err.Description
    IsIniValueLine = False
End Function

'' コンボボックスに2次元配列をセット
Public Sub FillComboBox(cmb As MSForms.ComboBox, arr As Variant)
    cmb.Clear
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        cmb.AddItem
        cmb.List(cmb.ListCount - 1, 0) = arr(i)(0)
        cmb.List(cmb.ListCount - 1, 1) = arr(i)(1)
    Next
End Sub

'' リストボックスに配列をセット
Public Sub FillListBox(lst As MSForms.ListBox, arr As Variant)
    lst.Clear
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        lst.AddItem arr(i)
    Next
End Sub

