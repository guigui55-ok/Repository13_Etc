Attribute VB_Name = "Module2"
'act_arr()
'act_setting()
'title_arr()
'num_arr()
'genre_arr()
'set_sell()


'/////////////////////////////////////////////////////////////////////////////////////////////////女優別並び替え
Public Function act_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    rect.stRow = rect.stCol = rect.Num = 0
    '出演女優　を一番左へ
    rect.cName = "出演女優"
    rect.Num = 1
    rect = LineCap(rect)
    'タイトル　をその右へ
    rect.cName = "タイトル"
    rect.Num = 2
    rect = LineCap(rect)
    '入力値すべて選択
    rect.cName = "出演女優"
    rect = allSelect(rect)
    '始点 + 1 から終点を選択
    If ActiveSheet.Name = "検索" Then
        rect.stRow = Worksheets("設定").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '並び替え
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), key2:=Columns(rect.stCol + 1)
    
    '元に戻す
    '出演女優　を4番目へ
    rect.cName = "タイトル"
    rect.Num = 6
    rect = LineCap(rect)
    'タイトル　をその右へ
    rect.cName = "出演女優"
    rect.Num = 6
    rect = LineCap(rect)
    
    Cells(1, 1).Select
    'ボタンのサイズと位置をセット
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////整列並び替え
Public Function act_setting()
    Dim rect2 As SellRect
    Dim cateCnt As Integer
    Dim cnt As Integer
    Dim tmpRange As Range
    Dim sName As String
    Dim setcnt As Integer
    Dim tmpObj As Object
    sName = ActiveSheet.Name
    
    '項目ワード数設定
    With Worksheets("設定")
        rect2.EndRow = .Cells(65535, 1).End(xlUp).Row
        rect2.EndCol = .Cells(65535, 1).End(xlUp).Column
        rect2.stRow = 1
        rect2.stCol = 1
        Set tmpObj = Range(.Cells(rect2.stRow, rect2.stCol), .Cells(rect2.EndRow, rect2.EndCol))
        cateCnt = tmpObj.Rows.Count
    End With
    
    Dim cateWord() As String
    ReDim cateWord(cateCnt + 2)
    '項目ワード読み込み
    With Worksheets("設定")
        cnt = 1
        'Set tmpRange = Selection
        For Each ForRange In tmpObj
            cateWord(cnt) = ForRange.Value
            cnt = cnt + 1
        Next ForRange
    End With
    '並び替え
    cnt = 1
    Worksheets(sName).Activate
    For cnt = 1 To cateCnt
        rect2.cName = cateWord(cnt)
        rect2.Num = cnt
        rect2 = LineCap(rect2)
    Next cnt
    
    Cells(1, 1).Select
    'ボタンのサイズと位置をセット
    Dim j As Integer
    Dim obj As OLEObject
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////タイトル別並び替え
Public Function title_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    'タイトル　をその右へ
    rect.cName = "タイトル"
    rect.Num = 1
    rect = LineCap(rect)
    '入力値すべて選択
    rect.cName = "タイトル"
    rect = allSelect(rect)
    '始点 + 1 から終点を選択
    If ActiveSheet.Name = "検索" Then
        rect.stRow = Worksheets("設定").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '並び替え　key2＝女優名
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), key2:=Columns(rect.stCol + 4)
    
    '元に戻す
    rect.cName = "タイトル"
    rect.Num = 5
    rect = LineCap(rect)
    Cells(Worksheets("設定").Cells(1, 4).Value, 1).Select
    Cells(1, 1).Select
'ボタンのサイズと位置をセット
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////番号順並び替え
Public Function num_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    'タイトル　をその右へ
    rect.cName = "No"
    rect.Num = 1
    rect = LineCap(rect)
    '入力値すべて選択
    rect.cName = "No"
    rect = allSelect(rect)
    '始点 + 1 から終点を選択
    If ActiveSheet.Name = "検索" Then
        rect.stRow = Worksheets("設定").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '並び替え　key2＝女優名
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), DataOption1:=xlSortTextAsNumbers
    
    '元に戻す
    rect.cName = "No"
    rect.Num = 1
    rect = LineCap(rect)
    Cells(Worksheets("設定").Cells(1, 4).Value, 1).Select
    Cells(1, 1).Select
'ボタンのサイズと位置をセット
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function


'/////////////////////////////////////////////////////////////////////////////////////////////////ジャンル順並び替え
Public Function genre_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    rect.stRow = rect.stCol = rect.Num = 0
    '出演女優　を一番左へ
    rect.cName = "ｼﾞｬﾝﾙ"
    rect.Num = 1
    rect = LineCap(rect)
'    'タイトル　をその右へ
'    rect.cName = "タイトル"
'    rect.num = 2
'    rect = LineCap(rect)
    '入力値すべて選択
    rect.cName = "ｼﾞｬﾝﾙ"
    rect = allSelect(rect)
    '始点 + 1 から終点を選択
    If ActiveSheet.Name = "検索" Then
        rect.stRow = Worksheets("設定").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '並び替え
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), key2:=Columns(rect.stCol + 1)
    
    '元に戻す
    rect.cName = "ｼﾞｬﾝﾙ"
    rect.Num = 8
    rect = LineCap(rect)
    Cells(Worksheets("設定").Cells(1, 4).Value, 1).Select
    Cells(1, 1).Select
    
    'ボタンのサイズと位置をセット
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
    
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////幅設定 set_sell
Public Function set_sell()
    Dim rect2 As SellRect
    Dim cateCnt As Integer
    Dim cnt As Integer
    Dim tmpRange As Range
    Dim sName As String
    Dim setcnt As Integer
    Dim tmpObj As Object

    '現在のシート名を記憶
    sName = ActiveSheet.Name
    
    '項目数
    cateCnt = Worksheets("設定").Cells(65535, 1).End(xlUp).Row
    '検索用項目名　cateword　宣言・確保
    Dim cateWord() As String
    ReDim cateWord(cateCnt + 2)
    '項目リストを配列へ
    cateWord = word_set()
    
    '並び替え
    cnt = 1
    Worksheets(sName).Activate
    For cnt = 1 To cateCnt
        rect2.cName = cateWord(cnt)
        rect2.Num = cnt
        rect2 = LineCap(rect2)
    Next cnt

    'rect2.cName = cateWord(2)
    'rect2 = allSelect2(rect2)
    
    '////////////////////////////////////////////////////////////////////
    '幅設定    //////////////////////////////////////////////////////////
    '読み込み
    setcnt = Worksheets("設定").Cells(65535, 9).End(xlUp).Row  '設定の9列目の最後の行番号
    Dim setVal() As String
    ReDim setVal(setcnt + 1)
    '項目ワード読み込み
    With Worksheets("設定")
        cnt = 1
        '選択
        Set tmpRange = .Range(.Cells(1, 9), .Cells(setcnt, 9))
        For Each ForRange In tmpRange
            setVal(cnt) = ForRange.Value
            cnt = cnt + 1
        Next ForRange
    End With
    
    'rows縦幅設定
    rect2.cName = cateWord(1) ' タイトル
    rect2 = find_set(rect2)
    rect2.EndRow = Cells(65535, rect2.stCol).End(xlUp).Row
    Set tmpObj = Range(Cells(rect2.stRow, rect2.stCol), Cells(rect2.EndRow, rect2.stCol))
    For Each ForRange In tmpObj
'        If forrange.RowHeight > 27 Or forrange.RowHeight < 10 Then
'            forrange.Rows.AutoFit
'        End If
        If ForRange.RowHeight <> 12 Then
            ForRange.RowHeight = 12
        End If
    Next ForRange
    'column幅設定
    Cells(1, 1).Select
    For cnt = 1 To setcnt
        Cells(1, cnt).ColumnWidth = setVal(cnt)
    Next cnt
'    '折り返して表示しない
'    For num = 6 To cateCnt + 1
'        rect2.cName = cateWord(num)
'        rect2 = find_set(rect2)
'        Set tmpObj = Cells(1, num).EntireColumn
'        If tmpObj.WrapText = True Then
'            tmpObj.WrapText = False
'        End If
'    Next num
'ボタンのサイズと位置をセット
    Dim j As Integer
    Dim obj As OLEObject
For Each obj In ActiveSheet.OLEObjects
    If obj.Name Like "CommandButton*" Then
        obj.Top = 0
        obj.Left = 300 + j * 70
        obj.Height = 23
        obj.Width = 50
        j = j + 1
    End If
Next

End Function
