Attribute VB_Name = "Bas_Update"
Option Explicit

'Public Sub subUpdate()
'    Dim ST          As Object:  Set ST = stList
'    Dim CN          As Object:  Set CN = Bas_DbConnection.GetConnection()
'    Dim strSQL      As String
'    Dim lResult     As Long
'    Dim lRow        As Long
'    Dim dltFLG      As Boolean: dltFLG = False
'    Dim i           As Long
'    Dim MaxRow      As Long
'    Dim sWorkKTCD   As String
'    Dim sEditKTCD   As String
'        
'    '最大行数取得
'    MaxRow = fncGetMaxRow
'    
'    'ＤＢ接続
'    ' DB接続は共通関数で取得済み
'
'    '================================================================
'    ' 1. 削除処理 (論理削除)
'    '================================================================
'    lRow = RW_FR
'    Do While Not stWork.Cells(lRow, 1) = ""
'        dltFLG = True
'        sWorkKTCD = fncGetKTCD(Trim(stWork.Cells(lRow, 1))) '退避データの工程コード
'        
'        For i = RW_FR To MaxRow
'            sEditKTCD = fncGetKTCD(Trim(ST.Cells(i, 1))) '編集データの工程コード
'            
'            '★Noだけでなく「工程」も一致しているか確認する
'            If sWorkKTCD = sEditKTCD And _
'               Val(stWork.Cells(lRow, 2)) = Val(StrConv(Trim(ST.Cells(i, 2)), vbNarrow)) Then
'                dltFLG = False
'                Exit For

'            End If

'        Next
