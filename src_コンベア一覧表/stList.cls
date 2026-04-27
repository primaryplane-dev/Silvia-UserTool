VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'èåèëIë
Public Sub ShowMenu()
    frmMenu.Show
End Sub

''ï¬Ç∂ÇÈ
'Private Sub cmdEnd_Click()
'    Me.Cells(1, 1).Copy
'    Application.CutCopyMode = False
'    Application.ThisWorkbook.Close False
'End Sub
'
'Private Sub cmdSelect_Click()
'    Call subOpen
'    If P_Regist Then
'        Call subMain
'    End If
'End Sub
'
'Private Sub subOpen()
'    Dim obj As New frmSelect
'    obj.Show
'    Set obj = Nothing
'End Sub
