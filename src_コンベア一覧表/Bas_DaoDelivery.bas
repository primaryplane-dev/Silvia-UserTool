Attribute VB_Name = "Bas_DaoDelivery"
Option Explicit

'' 品名リスト取得（SBGP01テーブル）
Public Function GetHinmList(ByVal yyyymmdd As String) As Object
    On Error GoTo ErrorHandler
    Dim objResult As Object
    Set objResult = CreateObject("ADODB.Recordset")
    Dim CN As Object: Set CN = Bas_DbConnection.GetConnection()
    Dim strSQL As String
    strSQL = "SELECT DISTINCT BGHINM, BGHINO, BGKJNO FROM LIBSMF17.SBGP01 " & _
             "WHERE BGDELT = '' AND BGSDAT = " & yyyymmdd & " ORDER BY BGHINO, BGKJNO"
    objResult.Open strSQL, CN, 3, 1 ' adOpenStatic=3, adLockReadOnly=1
    Set GetHinmList = objResult
    Exit Function
ErrorHandler:
    Set GetHinmList = Nothing
End Function

' コンベア一覧取得（SBFP01テーブル）
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
