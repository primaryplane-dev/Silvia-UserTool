VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSYCD 
   Caption         =   "担当者"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   OleObjectBlob   =   "frmSYCD.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSYCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'確定ボタン
Private Sub cmdOK_Click()
    Dim strSYCD     As String
    Dim strSYNM     As String

    '条件を保存する(ComboBox優先)
    If cbSYCD.Text <> "" Then
        strSYCD = Format(cbSYCD.Value, "00000000")
        strSYNM = cbSYCD.Text
    ElseIf lblSYNM.Caption <> "" Then
        strSYCD = Format(txtSYCD.Text, "00000000")
        strSYNM = lblSYNM.Caption
    End If
    If strSYCD = "" Then Exit Sub

    P_SYCD = strSYCD
    P_SYNM = strSYNM
    Unload Me
End Sub

'数値パッド
Private Sub cmdInputSYCD_Click()
    P_TenKeyData = txtSYCD.Text
    P_TenKeyKeta = 8:   P_TenKeyMinus = False:  P_TenKeyPointMode = 0
    frmTenKey.Show
    If P_TenKey_FLG Then
        txtSYCD.Text = P_TenKeyData
        lblSYNM.Caption = ""
        Call GetSYNM
    End If
End Sub

'社員マスタの存在チェック＆社員名を取得する
Private Sub GetSYNM()
    Dim strSYCD     As String:  strSYCD = txtSYCD.Text
    Dim strSYNM     As String

    '入力チェック
    If Val(strSYCD) = 0 Then Exit Sub
    
'    'Ini読込
'    If P_ConnectString = "" Then Call subGetIniFile

    '社員マスタを読む
    If Not GetSYCD(strSYCD, strSYNM) Then MsgBox ("社員コードが間違っています"): Exit Sub
    
    txtSYCD.Text = strSYCD
    lblSYNM.Caption = strSYNM

End Sub

Private Sub UserForm_Initialize()
    Call MakecbSYCD
End Sub

Public Sub MakecbSYCD()
    Dim i           As Integer
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    
'    'Ini読込
'    If P_ConnectString = "" Then Call subGetIniFile
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    'STSP01:   担当者マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT "
    strSQL = strSQL & " TSSYCD, TSSYKJ "
    strSQL = strSQL & " FROM  LIBSMF17.STSP01  "
    strSQL = strSQL & " WHERE TSDELT='' "
    strSQL = strSQL & " ORDER BY TSSYCD"
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    cbSYCD.Clear
    cbSYCD.AddItem
    Do While Not RS.EOF
        cbSYCD.AddItem
        cbSYCD.List(cbSYCD.ListCount - 1, 0) = Trim(RS("TSSYCD"))
        cbSYCD.List(cbSYCD.ListCount - 1, 1) = Trim(RS("TSSYKJ"))
        RS.MoveNext
    Loop
    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing
End Sub

Private Function GetSYCD(ByVal iSYCD As String, ByRef oSYNM As String) As Boolean
    Dim CN          As Object:  Set CN = CreateObject("ADODB.Connection")
    Dim RS          As Object:  Set RS = CreateObject("ADODB.Recordset")
    Dim strSQL      As String
    
    GetSYCD = False
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    

    'BVAP01：従業員管理マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT VASYCD, VASYKJ "
    strSQL = strSQL & " FROM  LIBBMF.BVAP01  "
    strSQL = strSQL & " WHERE VAKYUK='' "
    strSQL = strSQL & "   AND VASYCD= " & Val(iSYCD)
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    Do While Not RS.EOF
        GetSYCD = True
        oSYNM = RS("VASYKJ")
        Exit Do
    Loop

    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing
End Function


