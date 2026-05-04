Attribute VB_Name = "Bas_Public"
Option Explicit

'Public Const P_ConnectString  As String = "Provider=IBMDA400;Data Source=FUJIPAN;User ID=ODBC001;Password=FJPN2480;"
Public Const P_ConnectString  As String = "Provider=IBMDA400;Data Source=HONSHA;User ID=ODBC001;Password=FJPN2480;"


Public P_ChkKBN As Integer      ' 区分
Public P_DATE As Date           ' 処理日付
Public P_HINO As String         ' 日番号 (例: 1)
Public P_KJNO As String         ' ラインNo (例: 001)
Public P_HINM As String         ' 商品名


Public P_CalendarSelected        As Boolean
Public P_calDATE        As Date


' レイアウト設定
Public Const CL_SY As Long = 55       ' システム管理用列
Public Const RW_FR As Long = 37       ' データ開始行
Public Const ROW_TIME As Long = 31    ' 時間表示行
Public Const ROW_NAME As Long = 35    ' 担当者名表示行



