VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectSHBU 
   Caption         =   "条件選択"
   ClientHeight    =   3336
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6735
   OleObjectBlob   =   "frmSelectSHBU.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectSHBU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    P_Regist2 = False
    Call subMakeCombo
    If Not P_SHBU = "" Then
        cmbSHBU.Value = P_SHBU
    End If
End Sub

Private Sub subMakeCombo()
    Dim arrItems As Variant
    arrItems = Bas_LogicConveyor.GetSHBUList()
    Bas_Utilities.FillComboBox cmbSHBU, arrItems
End Sub

Private Sub cmdRegist_Click()
    If cmbSHBU.Text = "" Then Exit Sub
    P_SHBU = cmbSHBU.Value
    P_SHNM = cmbSHBU.Text
    P_Regist2 = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

