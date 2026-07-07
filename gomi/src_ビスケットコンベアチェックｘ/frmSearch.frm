VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "êªïiåüçı"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13035
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'ÉIÅ[ÉiÅ[ ÉtÉHÅ[ÉÄÇÃíÜâõ
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRegist_Click()
    If lstData.Text = "" Then Exit Sub
    P_Regist2 = True
    P_SelHNO = lstData.Value
    P_SelHNM = lstData.Text
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    P_Regist2 = False
    txtWord.Text = ""
    lstData.Clear
End Sub

Private Sub cmdSearch_Click()
    txtWord.Text = Trim(txtWord.Text)
    If txtWord.Text = "" Then Exit Sub
    Call subEditDataList
End Sub

Private Sub subEditDataList()
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    
    'ÇcÇaê⁄ë±
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT RHHNO, HMHNM"
    strSQL = strSQL & "   FROM LIBSMF17.SRHP01 "
    strSQL = strSQL & "   LEFT JOIN LIBWMF.WHMP01 "
    strSQL = strSQL & "     ON HMDLT = '' "
    strSQL = strSQL & "    AND HMHNO = RHHNO "
    strSQL = strSQL & "  WHERE (" & fncEditWhere(txtWord.Text) & ")"
    strSQL = strSQL & "    AND RHDLT = '' "
    
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    
    lstData.Clear
    Do While Not RS.EOF
        lstData.AddItem
        lstData.List(lstData.ListCount - 1, 0) = RS("RHHNO")
        lstData.List(lstData.ListCount - 1, 1) = RS("HMHNM")
        RS.MoveNext
    Loop
    
    'ÇcÇaêÿíf
    RS.Close:   Set RS = Nothing
    CN.Close:   Set CN = Nothing

End Sub

Private Function fncEditWhere(ByVal Word As String) As String
    Dim sSp()       As String: sSp = Split(Replace(Word, "Å@", " "), " ")
    Dim sSpZ()      As String: ReDim sSpZ(UBound(sSp))
    Dim sSpH()      As String: ReDim sSpH(UBound(sSp))
    Dim i           As Long
    Dim str         As String
    
    For i = 0 To UBound(sSp)
        sSpZ(i) = StrConv(sSp(i), vbWide)
        sSpH(i) = StrConv(sSp(i), vbNarrow)
    Next
    fncEditWhere = ""
    For i = 0 To UBound(sSp)
        If Not fncEditWhere = "" Then fncEditWhere = fncEditWhere & " AND "
        str = ""
        str = str & "    HMKNM LIKE '%" & sSp(i) & "%' OR HMHNM LIKE '%" & sSp(i) & "%'"
        str = str & " OR HMKNM LIKE '%" & sSpZ(i) & "%' OR HMHNM LIKE '%" & sSpZ(i) & "%'"
        str = str & " OR HMKNM LIKE '%" & sSpH(i) & "%' OR HMHNM LIKE '%" & sSpH(i) & "%'"
        fncEditWhere = fncEditWhere & "(" & str & ")"
    Next
    
End Function
