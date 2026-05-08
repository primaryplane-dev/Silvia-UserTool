Attribute VB_Name = "Bas_DaoConveyor"
Option Explicit

'/**
' * @file Bas_DaoConveyor.bas
' * @brief コンベアデータアクセス層 (DAL)
' * @note コンベア情報の取得および更新を行うための実装
' */

'' コンベア一覧取得（SBFP01テーブル）
Public Function GetConveyorList() As Object
    On Error GoTo ErrorHandler
    Dim objResult As Object
    Set objResult = CreateObject("ADODB.Recordset")
    Dim CN As Object: Set CN = Bas_DbConnection.GetConnection()
    Dim strSQL As String
    strSQL = "SELECT BFKTCD, BFCVNO, BFCVNM FROM LIBSMF17.SBFP01 WHERE BFDELT='' ORDER BY BFKTCD, BFCVNO"
    objResult.Open strSQL, CN, 3, 1 ' adOpenStatic=3, adLockReadOnly=1
    Set GetConveyorList = objResult
    Exit Function
ErrorHandler:
    Set GetConveyorList = Nothing
End Function

'/**
' * @brief コンベア一覧の更新・新規登録・論理削除
' * @param stList 編集後データ（シート等のリストオブジェクト）
' * @param stWork 編集前データ（シート等のリストオブジェクト）
' * @param RW_FR データ開始行
' * @note 旧subUpdateのDAO化。呼び出し元で引数を渡すこと。
' */
Public Sub UpdateConveyorList(ByVal stList As Object, ByVal stWork As Object, ByVal RW_FR As Long)
    On Error GoTo ErrorHandler
    Dim CN As Object: Set CN = Bas_DbConnection.GetConnection()
    Dim strSQL As String
    Dim lResult As Long
    Dim lRow As Long
    Dim dltFLG As Boolean
    Dim i As Long
    Dim MaxRow As Long
    Dim sWorkKTCD As String
    Dim sEditKTCD As String

    '最大行数取得
    MaxRow = fncGetMaxRow()

    '--- 1. 論理削除 ---
    lRow = RW_FR
    Do While Not stWork.Cells(lRow, 1) = ""
        dltFLG = True
        sWorkKTCD = fncGetKTCD(Trim(stWork.Cells(lRow, 1)))
        For i = RW_FR To MaxRow
            sEditKTCD = fncGetKTCD(Trim(stList.Cells(i, 1)))
            If sWorkKTCD = sEditKTCD And _
               Val(stWork.Cells(lRow, 2)) = Val(StrConv(Trim(stList.Cells(i, 2)), vbNarrow)) Then
                dltFLG = False
                Exit For
            End If
        Next
        If dltFLG Then
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBSMF17.SBFP01 SET"
            strSQL = strSQL & "     BFUNTH = TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & ",BFUTIM = TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & ",BFDELT = 'X'"
            strSQL = strSQL & " WHERE BFDELT='' "
            strSQL = strSQL & "   AND BFCVNO=" & Val(stWork.Cells(lRow, 2))
            strSQL = strSQL & "   AND BFKTCD='" & sWorkKTCD & "'"
            CN.Execute strSQL, lResult, &H80
        End If
        lRow = lRow + 1
    Loop

    '--- 2. 更新・新規登録 ---
    For lRow = RW_FR To MaxRow
        If Trim(stList.Cells(lRow, 1)) <> "" And Val(StrConv(Trim(stList.Cells(lRow, 2)), vbNarrow)) > 0 And Trim(stList.Cells(lRow, 3)) <> "" Then
            sEditKTCD = fncGetKTCD(Trim(CStr(stList.Cells(lRow, 1))))
            '--- UPDATE ---
            strSQL = ""
            strSQL = strSQL & "UPDATE LIBSMF17.SBFP01 SET"
            strSQL = strSQL & "     BFUNTH =TO_CHAR(current timestamp, 'YYYYMMDD')"
            strSQL = strSQL & ",BFUTIM =TO_CHAR(current timestamp, 'HH24MISS')"
            strSQL = strSQL & ",BFCVNM = '" & Trim(CStr(stList.Cells(lRow, 3))) & "'"
            strSQL = strSQL & ",BFKTCD = '" & sEditKTCD & "'"
            strSQL = strSQL & " WHERE BFDELT='' "
            strSQL = strSQL & "   AND BFCVNO=" & Val(StrConv(Trim(stList.Cells(lRow, 2)), vbNarrow))
            strSQL = strSQL & "   AND BFKTCD='" & sEditKTCD & "'"
            CN.Execute strSQL, lResult, &H80
            '--- INSERT ---
            If lResult = 0 Then
                strSQL = ""
                strSQL = strSQL & " INSERT INTO LIBSMF17.SBFP01"
                strSQL = strSQL & "         (BFDELT,BFCPGM,BFCNTH,BFCTIM,BFKTCD,BFCVNO,BFCVNM) VALUES ("
                strSQL = strSQL & "              ''"
                strSQL = strSQL & ", '" & SYSTEM_NAME & "'"
                strSQL = strSQL & ", TO_CHAR(current timestamp, 'YYYYMMDD')"
                strSQL = strSQL & ", TO_CHAR(current timestamp, 'HH24MISS')"
                strSQL = strSQL & ", '" & sEditKTCD & "'"
                strSQL = strSQL & "," & Val(StrConv(Trim(stList.Cells(lRow, 2)), vbNarrow))
                strSQL = strSQL & ", '" & Trim(CStr(stList.Cells(lRow, 3))) & "'"
                strSQL = strSQL & ")"
                CN.Execute strSQL, lResult, &H80
            End If
        End If
    Next

    CN.Close: Set CN = Nothing
    Exit Sub
ErrorHandler:
    If Not CN Is Nothing Then CN.Close: Set CN = Nothing
    Call MsgBox("コンベア情報の更新処理でエラーが発生しました。管理者へ連絡してください。", vbCritical, SYSTEM_NAME)
End Sub
