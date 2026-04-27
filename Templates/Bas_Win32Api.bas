Attribute VB_Name = "Bas_Win32Api"
'/**
' * @file Bas_Win32Api.bas
' * @brief Win32API 宣言モジュール
' * @note 32-bit / 64-bit 両環境に対応した PtrSafe 宣言を行う
' */

Option Explicit

'/**
' * @brief 指定したミリ秒の間、処理を一時停止する
' * @param dwMilliseconds 停止する時間 (ミリ秒)
' */
#If VBA7 Then
    ' 64-bit (Office 2010以降)
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    ' 32-bit (Legacy)
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
