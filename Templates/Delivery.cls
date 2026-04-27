VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * @file Delivery.cls
' * @brief 【SAMPLE】配送データ格納用 DTO (Data Transfer Object)
' * @note データベースの 1 レコード分、または画面の入力データを保持する
' */

Option Explicit

'/** @brief 配送 ID */
Public DeliveryID As Long

'/** @brief 配送先名称 */
Public DestinationName As String

'/** @brief 配送予定日 */
Public DeliveryDate As Date

'/** @brief 配送ステータス (例: "手配中", "完了", "キャンセル") */
Public Status As String

'/**
' * @brief クラスの初期化処理
' * @note 各フィールドに安全な初期値を設定する
' */
Private Sub Class_Initialize()
    DeliveryID = 0
    DestinationName = ""
    DeliveryDate = CDate("1900/01/01")
    Status = ""
End Sub
