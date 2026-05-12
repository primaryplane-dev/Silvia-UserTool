VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelect 
   Caption         =   "条件指定"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   OleObjectBlob   =   "frmSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
'    If Not P_DATEF = 0 Then
'        txtDateF.Text = Format(P_DATEF, "yyyy/mm/dd")
'    End If
'    If Not P_DATET = 0 Then
'        txtDateT.Text = Format(P_DATET, "yyyy/mm/dd")
'    End If
    P_Regist = False
End Sub

Private Sub cmdRegist_Click()
    If txtDateF.Text = "" Then
        P_DATEF = 0
    Else
        P_DATEF = CDate(txtDateF.Text)
    End If
    If txtDateT.Text = "" Then
        P_DATET = Date
    Else
        P_DATET = CDate(txtDateT.Text)
    End If
    P_Regist = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCalendar1_Click()
    Call subOpenCalendar
    If Not P_Regist2 Then Exit Sub
    txtDateF.Text = Format(P_calDATE, "yyyy/mm/dd")
End Sub

Private Sub cmdCalendar2_Click()
    Call subOpenCalendar
    If Not P_Regist2 Then Exit Sub
    txtDateT.Text = Format(P_calDATE, "yyyy/mm/dd")
End Sub

Private Sub subOpenCalendar()
    Dim obj As New frmCalendar
    obj.Show
    Set obj = Nothing
End Sub
