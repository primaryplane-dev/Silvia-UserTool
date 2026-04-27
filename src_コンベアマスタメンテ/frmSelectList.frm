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
    P_SelItem = lstItem.List(lstItem.ListIndex)
    P_Regist = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Call subMakeList
    P_SelItem = ""
    P_Regist = False
End Sub

Private Sub subMakeList()
    
    lstItem.Clear
    
    ' 空白行
    lstItem.AddItem
    lstItem.List(lstItem.ListCount - 1, 0) = ""
    
    ' 指定の工程のみ追加
    lstItem.AddItem
    lstItem.List(lstItem.ListCount - 1, 0) = "成型"
    
    lstItem.AddItem
    lstItem.List(lstItem.ListCount - 1, 0) = "冷却"
    
End Sub
