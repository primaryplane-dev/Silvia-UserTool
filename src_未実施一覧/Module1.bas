Attribute VB_Name = "Module1"
Option Explicit

Public Sub ExportAllModules()
    ' ※注意：事前準備の「VBAプロジェクトへのアクセスを信頼する」がオンでないとエラーになる
    
    Dim cmp As Object
    Dim exportPath As String
    Dim ext As String
    
    ' エクスポート先のフォルダパス（このExcelファイルと同じ階層の "src" フォルダとする場合）
    exportPath = ThisWorkbook.Path & "\src\"
    
    ' 保存先のフォルダが存在しない場合は作成する
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' 全てのモジュールをループしてエクスポート
    For Each cmp In ThisWorkbook.VBProject.VBComponents
        ' モジュールの種類によって拡張子を決定
        Select Case cmp.Type
            Case 1 ' 標準モジュール
                ext = ".bas"
            Case 2 ' クラスモジュール
                ext = ".cls"
            Case 3 ' ユーザーフォーム
                ext = ".frm"
            Case 100 ' ThisWorkbookやSheetなどのドキュメントモジュール
                ext = ".cls"
            Case Else
                ext = ".txt"
        End Select
        
        ' エクスポート実行
        cmp.Export exportPath & cmp.Name & ext
    Next cmp
    
    ' 実行完了をステータスバーに表示（メッセージボックスだと毎回表示されて煩わしいため）
    Application.StatusBar = "VBAコードのエクスポートが完了しました (" & Now & ")"
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar" ' 3秒後にクリア
End Sub

' ステータスバーをクリアする用のサブプロシージャ
Private Sub ClearStatusBar()
    Application.StatusBar = False
End Sub
