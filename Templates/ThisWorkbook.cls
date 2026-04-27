VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'/**
' * @file ThisWorkbook.cls
' * @brief ブック全体のイベントを管理するクラスモジュール
' * @note Workbook_Open など、ブックライフサイクルに関する処理のみを記述する。
' *       業務ロジックは BLL 層へ委譲し、本モジュールには記述しない。
' */

Option Explicit
