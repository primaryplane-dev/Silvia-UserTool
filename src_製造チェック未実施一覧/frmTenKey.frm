VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTenKey 
   Caption         =   "数値パッド"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   OleObjectBlob   =   "frmTenKey.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTenKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private ClearFLG    As Boolean

Private Sub btnOk_Click()
    Dim strMSG      As String
    
    '入力チェック
    If Not ClearFLG Then
        If lblInputValue.Caption = "" Then Exit Sub
    End If
    strMSG = CheckKeta(lblInputValue.Caption)
    If strMSG <> "" Then Call MsgBox(strMSG): Exit Sub
    
   
    P_TenKeyData = lblInputValue.Caption
    P_TenKey_FLG = True
    Unload Me
End Sub

Private Function CheckKeta(ByVal i_Val As String) As String
    CheckKeta = ""
    If i_Val = "" Then Exit Function
    'マイナスは桁に入れない
    Dim strWK   As String:  strWK = Replace(i_Val, "-", "")

    If P_TenKeyPointMode = 0 Then
        '正の整数のみの場合
        If Len(strWK) > P_TenKeyKeta Then CheckKeta = P_TenKeyKeta & "桁以内で入力してください"
        Exit Function
    End If
    
    Dim varWK   As Variant: varWK = Split(strWK, ".")
    '整数部の桁チェック/小数部は１桁のみ(今のところ)
    If Len(varWK(0)) > P_TenKeyKeta Then CheckKeta = "整数部は" & P_TenKeyKeta & "桁以内で入力してください"
    If UBound(varWK) >= 1 Then
        If P_TenKeyPointMode = 1 Then
            If Len(varWK(1)) > 1 Then CheckKeta = "小数部は１桁以内で入力してください"
        ElseIf P_TenKeyPointMode = 2 Then
            If Len(varWK(1)) > 2 Then CheckKeta = "小数部は２桁以内で入力してください"
        End If
    End If
End Function

Private Sub UserForm_Initialize()
    lblInputValue.Caption = P_TenKeyData
    P_TenKey_FLG = False
    P_TEdit_FLG = False
    ClearFLG = False
    
    'マイナス、小数点ボタンの表示切替
    If Not P_TenKeyPointMode = 0 And P_TenKeyMinus Then
        btnD.Visible = True
        btnMinus.Visible = True
    ElseIf Not P_TenKeyPointMode = 0 And Not P_TenKeyMinus Then
        btnMinus.Visible = False
        btnD.Visible = True
        btnC.Width = 78
        btnD.Width = 78:   btnD.Left = 84
        btn0.Width = 78:   btn0.Left = 162
    Else
        btnMinus.Visible = False
        btnD.Visible = False
        btnC.Width = 78
        btn0.Width = 156:   btn0.Left = 84
    End If
'    Me.StartUpPosition = 0
'    Me.Left = 640:  Me.Top = 50
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub subTenKeyClick(ByVal key As String)
    Dim strLabel As String
    
    If Not P_TEdit_FLG Then
        strLabel = ""
        P_TEdit_FLG = True
    Else
        strLabel = lblInputValue.Caption
    End If
    Select Case key
    Case "C"
        strLabel = "":  ClearFLG = True
    Case "-"
        If InStr(strLabel, "-") > 0 Then
            strLabel = Replace(strLabel, "-", "")
        Else
            strLabel = "-" & strLabel
        End If
    Case "."
'        If InStr(1, strLabel, ".") = 0 Then
        Select Case strLabel
            Case ""
                strLabel = "0."
            Case Else
                strLabel = strLabel & "."
        End Select
'        End If
    Case Else
        strLabel = strLabel & key
    End Select
    
    lblInputValue.Caption = strLabel

End Sub

'------------------------------------------------------------------------------------------------
' 数値キー
'------------------------------------------------------------------------------------------------
Private Sub btnD_Click()
    Call subTenKeyClick(".")
End Sub

Private Sub btnMinus_Click()
    Call subTenKeyClick("-")
End Sub

Private Sub btnC_Click()
    Call subTenKeyClick("C")
End Sub

Private Sub btn0_Click()
    Call subTenKeyClick("0")
End Sub

Private Sub btn0_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("0")
End Sub

Private Sub btn1_Click()
    Call subTenKeyClick("1")
End Sub
Private Sub btn1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("1")
End Sub

Private Sub btn2_Click()
    Call subTenKeyClick("2")
End Sub
Private Sub btn2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("2")
End Sub

Private Sub btn3_Click()
    Call subTenKeyClick("3")
End Sub
Private Sub btn3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("3")
End Sub

Private Sub btn4_Click()
    Call subTenKeyClick("4")
End Sub
Private Sub btn4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("4")
End Sub

Private Sub btn5_Click()
    Call subTenKeyClick("5")
End Sub
Private Sub btn5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("5")
End Sub

Private Sub btn6_Click()
    Call subTenKeyClick("6")
End Sub
Private Sub btn6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("6")
End Sub

Private Sub btn7_Click()
    Call subTenKeyClick("7")
End Sub
Private Sub btn7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("7")
End Sub

Private Sub btn8_Click()
    Call subTenKeyClick("8")
End Sub
Private Sub btn8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("8")
End Sub

Private Sub btn9_Click()
    Call subTenKeyClick("9")
End Sub
Private Sub btn9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call subTenKeyClick("9")
End Sub


