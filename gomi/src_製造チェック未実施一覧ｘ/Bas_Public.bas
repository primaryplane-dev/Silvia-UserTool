Attribute VB_Name = "Bas_Public"
Option Explicit

'Public Const P_ConnectString  As String = "Provider=IBMDA400;Data Source=FUJIPAN;User ID=ODBC001;Password=FJPN2480;"
Public Const P_ConnectString  As String = "Provider=IBMDA400;Data Source=HONSHA;User ID=ODBC001;Password=FJPN2480;"

'frmSelectÇ©ÇÁ
Public P_DATEF              As Date
Public P_DATET              As Date
Public P_Regist             As Boolean
'frmCalendar
Public P_Regist2            As Boolean
Public P_calDATE            As Date
'frmSYCD
Public P_SYCD               As String
Public P_SYNM               As String
'frmTenKey
Public P_TenKeyData         As String
Public P_TenKeyKeta         As Integer
Public P_TenKeyMinus        As Boolean  'ïâ
Public P_TenKeyPointMode    As Integer  'è¨êîì_
Public P_TenKey_FLG         As Boolean
Public P_TEdit_FLG          As Boolean

Public P_DATE               As Date
Public P_HINO               As String
Public P_KJNO               As String
Public P_HINM               As String
Public P_SHBU               As String
Public P_KTCD               As String

Public Const P_PGM = "EXCEL"
