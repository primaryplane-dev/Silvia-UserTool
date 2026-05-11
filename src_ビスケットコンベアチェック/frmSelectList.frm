VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectList 
   Caption         =   "選択"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   OleObjectBlob   =   "frmSelectList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstItem_Click()
    P_SelCD = lstItem.List(lstItem.ListIndex, 0)
    P_SelNM = lstItem.List(lstItem.ListIndex, 1)
    P_Regist3 = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Call subMakeList
    P_SelCD = ""
    P_SelNM = ""
    P_Regist = False
End Sub

Private Sub subMakeList()
    Dim varKey      As Variant
    
    lstItem.Clear
    lstItem.AddItem
    lstItem.List(lstItem.ListCount - 1, 0) = ""
    lstItem.List(lstItem.ListCount - 1, 1) = ""
    If P_SelectKBN = 1 Then
        lstItem.AddItem
        lstItem.List(lstItem.ListCount - 1, 0) = "1"
        lstItem.List(lstItem.ListCount - 1, 1) = "〇"   '漢字の丸じゃないとoffice2016では表示がおかしくなる
        lstItem.AddItem
        lstItem.List(lstItem.ListCount - 1, 0) = "0"
        lstItem.List(lstItem.ListCount - 1, 1) = "×"
        lstItem.AddItem
        lstItem.List(lstItem.ListCount - 1, 0) = "2"
        lstItem.List(lstItem.ListCount - 1, 1) = "－"
    Else
        Dim CN      As New ADODB.Connection
        Dim RS      As New ADODB.Recordset
        Dim strSQL  As String
        
        'DB接続
        CN.CursorLocation = adUseClient
        CN.Open P_ConnectString
                
        strSQL = ""
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "   TSSYCD "
        strSQL = strSQL & "  ,TSSYKJ "
        strSQL = strSQL & "  FROM LIBSMF17.STSP01 "
        strSQL = strSQL & " WHERE TSDELT = '' "
        strSQL = strSQL & "   AND TSTTKB = '4'"
    
        RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
        
        Do While Not RS.EOF
            lstItem.AddItem
            lstItem.List(lstItem.ListCount - 1, 0) = RS("TSSYCD")
            lstItem.List(lstItem.ListCount - 1, 1) = RS("TSSYKJ")
            RS.MoveNext
        Loop
    
        'DB切断
        RS.Close: Set RS = Nothing
        CN.Close: Set CN = Nothing
    End If
End Sub
