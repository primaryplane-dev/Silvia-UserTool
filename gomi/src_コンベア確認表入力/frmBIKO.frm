VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBIKO 
   Caption         =   "備考"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10245
   OleObjectBlob   =   "frmBIKO.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmBIKO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    txtBIKO.Text = P_BIKO
    P_Regist4 = False
End Sub

Private Sub cmdRegist_Click()
    If LenAS(txtBIKO.Text) > 200 Then MsgBox ("備考は全角99桁以内"): Exit Sub
    P_BIKO = txtBIKO.Text

    P_Regist4 = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtBIKO.Text = ""
End Sub
