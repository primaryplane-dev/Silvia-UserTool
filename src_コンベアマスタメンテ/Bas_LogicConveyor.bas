Attribute VB_Name = "Bas_LogicConveyor"
'/**
' * @file Bas_LogicConveyor.bas
' * @brief コンベア業務ロジック層 (BLL)
' * @note コンベアに関する業務ルールを適用する実装
'
Option Explicit

'/**
' * @brief SHBU選択肢リストを取得
' * @return Variant SHBUリスト
'
Public Function GetSHBUList() As Variant
    GetSHBUList = Array(Array("", ""), Array("1", "ビスケット"), Array("2", "クッキー"), Array("3", "ドーナツ"))
End Function

'/**
' * @brief 工程リストを取得
' * @return Variant 工程リスト
'
Public Function GetSelectableProcessList() As Variant
    GetSelectableProcessList = Array("", "成型", "冷却")
End Function

'/**
' * @brief コンベア一覧の業務ロジック（例：名称のフィルタや加工など）
' * @return Object 加工済みRecordset
'
Public Function GetConveyorListWithLogic() As Object
    On Error GoTo ErrorHandler
    Dim rsRaw As Object: Set rsRaw = Bas_DaoConveyor.GetConveyorList()
    Dim rsResult As Object: Set rsResult = CreateObject("ADODB.Recordset")
    ' フィールド定義
    rsResult.Fields.Append "工程", 200, 32 ' AD_VAR_CHAR = 200
    rsResult.Fields.Append "No", 3 ' AD_INTEGER = 3
    rsResult.Fields.Append "名称", 200, 64
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
    Set GetConveyorListWithLogic = Nothing
End Function
