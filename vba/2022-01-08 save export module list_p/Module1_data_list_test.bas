Attribute VB_Name = "Module1"
'LineCap
'allSelect
'word_set()


'///////////////////////////////////////////////////////////////////////////// lineCap
'このようにして使う
'    rect.cName = "タイトル"
'    rect.num = 1
'    rect = LineCap(rect)
Public Function LineCap(ByRef lineRect As SellRect) As SellRect
    Num = lineRect.Num

    'カテゴリ名を探す
    lineRect = find_set(lineRect)
    'find flag がゼロなら抜ける
    If lineRect.Fflag = 0 Then
        Exit Function
    End If
    'あったら
    Cells(lineRect.stRow, lineRect.stCol).EntireColumn.Select
    
    If (Selection.Column) = (lineRect.Num) Then
        'Cells(1, 1).Select
    Else
        Selection.Cut
        Cells(1, Num).EntireColumn.Insert Shift:=xlShiftToRight
    End If
    LineCap = lineRect
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////// allSelect
' rect.cName = "タイトル"
' rect = allSelect(rect)
Public Function allSelect(ByRef asRect As SellRect) As SellRect
    Dim rowcnt As Integer
    asRect = find_set(asRect)
    'Cells(1, 1).Value = asRect.cName
    '行の終わりを選択
'    Cells(65535, 1).End(xlUp).Select
'    asRect.endRow = Selection.Row
    asRect.EndRow = Cells(65535, 1).End(xlUp).Row
    '列の終わりを選択
    'Cells(asRect.stRow, 100).End(xlToLeft).Select
    'asRect.endCol = Selection.Column
    asRect.EndCol = Cells(asRect.stRow, 100).End(xlToLeft).Column
    '始点 + 1 から終点を選択
    If ActiveSheet.Name = "検索" Then
        asRect.stRow = Worksheets("設定").Cells(1, 4).Value
        Range(Cells(asRect.stRow + 1, asRect.stCol), Cells(asRect.EndRow, asRect.EndCol)).Select
    Else
        Range(Cells(asRect.stRow + 1, asRect.stCol), Cells(asRect.EndRow, asRect.EndCol)).Select
    End If

    allSelect = asRect
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////// allSelect2
' rect.cName = "タイトル"
' rect = allSelect(rect)
Public Function allSelect2(ByRef asRect As SellRect) As SellRect
    
    Dim rowcnt As Integer
    Dim Num As Integer
    rowcnt = Worksheets("設定").Cells(65535, 1).End(xlUp).Row
    asRect = find_set(asRect)
    'Cells(1, 1).Value = asRect.cName
    '行の終わりを選択
'    Cells(65535, 1).End(xlUp).Select
'    asRect.endRow = Selection.Row
    asRect.EndRow = Cells(65535, 1).End(xlUp).Row
    'すべての終わりの行を調査
    For Num = 1 To rowcnt
        If Cells(65535, Num).End(xlUp).Row > asRect.EndRow Then
            asRect.EndRow = CInt(Cells(65535, Num).End(xlUp).Row)
        End If
    Next Num
    '列の終わりを選択
    'Cells(asRect.stRow, 100).End(xlToLeft).Select
    'asRect.endCol = Selection.Column
    asRect.EndCol = CInt(Cells(asRect.stRow, 100).End(xlToLeft).Column)
    '始点 + 1 から終点を選択
    If ActiveSheet.Name = "検索" Then
        asRect.stRow = Worksheets("設定").Cells(1, 4).Value
        Range(Cells(asRect.stRow + 1, asRect.stCol), Cells(asRect.EndRow, asRect.EndCol)).Select
    Else
        Range(Cells(asRect.stRow + 1, asRect.stCol), Cells(asRect.EndRow, asRect.EndCol)).Select
    End If

    allSelect2 = asRect
End Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////word_set
Public Function word_set()
    Dim rect2 As SellRect
    Dim cateCnt As Integer
    Dim cnt As Integer
    Dim tmpRange As Range
    Dim sName As String
    Dim setcnt As Integer
    Dim tmpObj As Object
    '一列目の最初から最後までの座標を　tmpObjへ
    With Worksheets("設定")
        rect2.EndRow = .Cells(65535, 1).End(xlUp).Row
        rect2.EndCol = .Cells(65535, 1).End(xlUp).Column
        rect2.stRow = 1
        rect2.stCol = 1
        Set tmpObj = Range(.Cells(rect2.stRow, rect2.stCol), .Cells(rect2.EndRow, rect2.EndCol))
        cateCnt = tmpObj.Rows.Count
    End With

    Dim cateWord2() As String
    ReDim cateWord2(cateCnt + 2)
    
    'tmpObj の値を配列へ
    With Worksheets("設定")
        cnt = 1
        'Set tmpRange = Selection
        For Each ForRange In tmpObj
            cateWord2(cnt) = ForRange.Value
            cnt = cnt + 1
        Next ForRange
    End With

    word_set = cateWord2

End Function
