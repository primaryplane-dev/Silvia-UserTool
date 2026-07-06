Option Explicit

'' 接続先切替用
'' 運用ルール:
'' 1) 本番接続は P_ConnectString の1行だけを有効化する
'' 2) 切替時は Bas_List / Bas_DbConnection の双方がこの定数を参照していることを確認する
'' 3) 接続先変更後は「件数・時間列・製造終了表示」をチェックリストで確認する
'Public Const P_ConnectString    As String = "Provider=IBMDA400;Data Source=FUJIPAN;User ID=ODBC001;Password=FJPN2480;"
Public Const P_ConnectString     As String = "Provider=IBMDA400;Data Source=HONSHA;User ID=ODBC001;Password=FJPN2480;"
Public Const P_PGM              As String = "EXCEL"

' frmMenu用
Public P_ChkKBN                 As Integer      ' 区分
Public P_DATE                   As Date         ' 処理日付
Public P_HINO                   As String       ' 日番号 (例: 1)
Public P_KJNO                   As String       ' ラインNo (例: 001)
Public P_HINM                   As String       ' 商品名
Public P_Regist2                As Boolean      ' 登録処理2回目以降かどうかのフラグ
Public P_CalendarSelected       As Boolean
Public P_calDATE                As Date

' レイアウト設定
Public Const CL_SY              As Long = 55    ' システム管理用列
Public Const RW_FR              As Long = 37    ' データ開始行
Public Const ROW_TIME           As Long = 31    ' 時間表示行
Public Const ROW_NAME           As Long = 35    ' 担当者名表示行



