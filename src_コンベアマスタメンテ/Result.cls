VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/**
' * @file Result.cls
' * @brief 処理結果格納用 DTO (Data Transfer Object)
' */


'/** @brief 処理が成功したかどうか */
Public IsSuccess As Boolean

'/** @brief ユーザーへの通知メッセージやエラー詳細 */
Public Message As String

'/** @brief 戻り値となるデータ (プリミティブ型、オブジェクト、配列など) */
Public Data As Variant

'/**
' * @brief クラスの初期化処理
' * @note フィールドの初期値を設定する
' */
Private Sub Class_Initialize()
    IsSuccess = False
    Message = ""
    Data = Empty
End Sub
