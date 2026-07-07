Attribute VB_Name = "Bas_List"
Option Explicit

Public Const RW_FR = 5

Public Sub subMain()
    Call subBeforeEdit
    Call subEditList
    Call subAfterEdit
End Sub

Private Sub subInitialize()
    stHina.Cells.Copy stList.Cells
    stList.Select
End Sub

Public Sub subEditList()
    Dim ST          As Worksheet: Set ST = stList
    Dim RS          As Object: Set RS = Bas_LogicConveyor.GetConveyorListWithLogic()
    Dim lRow        As Long
    
    ' BLLで加工済みのため不要

    '雛型シート→編集シート（ヘッダ部のみコピー）
    stHina.Rows("1:" & RW_FR - 1).Copy ST.Rows(1)
    ' DB件数分だけひな型行をコピー
    Dim tmpRow As Long: tmpRow = RW_FR
    If Not RS Is Nothing Then
        Do While Not RS.EOF
            stHina.Rows(RW_FR).Copy ST.Rows(tmpRow)
            tmpRow = tmpRow + 1
            RS.MoveNext
        Loop
        RS.MoveFirst '値セット用にカーソルを先頭に戻す
    End If

    '見出し
    ST.Cells(2, 1) = "アイテム名：共通"

    'DB件数分だけ値をセット
    lRow = RW_FR
    If Not RS Is Nothing Then
        Do While Not RS.EOF
            ST.Cells(lRow, 1) = RS("工程")
            ST.Cells(lRow, 2) = RS("No")
            ST.Cells(lRow, 3) = RS("名称")
            lRow = lRow + 1
            RS.MoveNext
        Loop
    End If
    
    '罫線 (1～3列目まで)
    If lRow > RW_FR Then
        ST.Range(ST.Cells(RW_FR, 1), ST.Cells(lRow - 1, 3)).Borders.LineStyle = xlContinuous
    End If
    
    '編集シート→ワークシート
    ST.Cells.Copy stWork.Cells
    
    ST.Cells(RW_FR - 1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    'ＤＢ切断
    RS.Close:   Set RS = Nothing
    Set ST = Nothing
End Sub
