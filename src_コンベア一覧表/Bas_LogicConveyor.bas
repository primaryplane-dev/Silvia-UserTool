Attribute VB_Name = "Bas_LogicConveyor"
'/**
' * @file Bas_LogicConveyor.bas
' * @brief コンベア業務ロジック層 (BLL)
' * @note コンベアに関する業務ルールを適用する実装例
' */

Option Explicit

'/**
' * @brief 品名リスト取得（業務ロジック層）
' * @param yyyymmdd 日付文字列
' * @return Recordset オブジェクト（必要に応じて加工）
' */
Public Function GetHinmListWithLogic(ByVal yyyymmdd As String) As Object
    On Error GoTo ErrorHandler

    Dim rsRaw As Object
    Set rsRaw = Bas_DaoDelivery.GetHinmList(yyyymmdd)
    
    ' 必要ならここでフィルタや加工を追加
    Set GetHinmListWithLogic = rsRaw
    
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_LogicConveyor.GetHinmListWithLogic: " & Err.Number & " - " & Err.Description
    Set GetHinmListWithLogic = Nothing
End Function

'/**
' * @brief SHBU選択肢リスト (UI層用)
' * @return Variant 二次元配列（値, 表示名）
' */
Public Function GetSHBUList() As Variant
    GetSHBUList = Array(Array("", ""), Array("1", "ビスケット"), Array("2", "クッキー"), Array("3", "ドーナツ"))
End Function

'/**
' * @brief 工程リスト (UI層用)
' * @return Variant 配列（工程名）
' */
Public Function GetSelectableProcessList() As Variant
    GetSelectableProcessList = Array("", "成型", "冷却")
End Function

'/**
' * @brief コンベア一覧の業務ロジック（例：名称のフィルタや加工など）
' * @return Recordset オブジェクト
' */
Public Function GetConveyorListWithLogic() As Object
    On Error GoTo ErrorHandler

    Dim rsRaw As Object
    Set rsRaw = Bas_DaoDelivery.GetConveyorList()
    
    Dim rsResult As Object
    Set rsResult = CreateObject("ADODB.Recordset")

    ' フィールド定義
    rsResult.Fields.Append "工程", AD_VAR_CHAR, 32
    rsResult.Fields.Append "No", AD_INTEGER
    rsResult.Fields.Append "名称", AD_VAR_CHAR, 64
    rsResult.Open

    If Not rsRaw Is Nothing Then
        Do While Not rsRaw.EOF
            ' 例：名称が空欄でないものだけ返す
            If Trim(rsRaw("BFCVNM") & "") <> "" Then
                rsResult.AddNew
                rsResult("工程") = fncGetKTNM(CStr(rsRaw("BFKTCD") & ""))
                rsResult("No") = rsRaw("BFCVNO")
                rsResult("名称") = rsRaw("BFCVNM")
                rsResult.Update
            End If
            rsRaw.MoveNext
        Loop
    End If

    Set GetConveyorListWithLogic = rsResult

    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, Bas_Configuration.LOG_DATE_FORMAT) & "] [Error] Bas_LogicConveyor.GetConveyorListWithLogic: " & Err.Number & " - " & Err.Description
    Set GetConveyorListWithLogic = Nothing
End Function
