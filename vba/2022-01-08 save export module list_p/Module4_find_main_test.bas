Attribute VB_Name = "Module4"

Type actStruct
    SheetName As String
    cate As String      '項目ワード
    find As String      '検索ワード
    rcnt As String      'rowsカウンタ
    Ccnt As String      'Cateカウンタ
End Type



'find_main
'actFind

Type SellRect
    stRow As Integer
    stCol As Integer
    EndRow As Integer
    EndCol As Integer
    tmpRow As Integer
    tmpCol As Integer
    cName As String
    Fflag As Integer
    Num As Integer
    sName As String
End Type
'//////////////////////////////////////////////////////////////////////////////////////////////
Public Function find_main()
    Dim rowcnt As Integer   '行カウント
    Dim rect2 As SellRect   '検索用
    Dim tmpRange As Range   '検索用
    Dim tmpObj As Object    '一時
    
    '項目ワード数設定
    With Worksheets("設定")
        rowcnt = .Cells(1, 4).Value
        rect2.EndRow = .Cells(65535, 1).End(xlUp).Row
        rect2.EndCol = .Cells(65535, 1).End(xlUp).Column
        rect2.stRow = 1
        rect2.stCol = 1
        Set tmpObj = .Range(.Cells(rect2.stRow, rect2.stCol), .Cells(rect2.EndRow, rect2.EndCol))
        cateCnt = tmpObj.Rows.Count
    End With
    
    '項目ワード数　確保
    Dim cateWord() As String
    ReDim cateWord(cateCnt + 2)
    '項目ワード読み込み　配列へ
    cnt = 1
    For Each ForRange In tmpObj
        cateWord(cnt) = ForRange.Value
        cnt = cnt + 1
    Next ForRange


    '検索ワード宣言
    Dim findWord() As String
    ReDim findWord(cateCnt + 1)
    '検索ワードを配列へ
    cnt = 1
    Set tmpRange = Range(Cells(5, 1), Cells(5, 1 + cateCnt))
    For Each ForRange In tmpRange
        findWord(cnt) = ForRange.Value
        cnt = cnt + 1
    Next ForRange
    

    '検索ワード貼り付け　お試し　デバック用
    cnt = 1
    Set tmpRange = Range(Cells(7, 1), Cells(7, 1 + cateCnt))
    For Each ForRange In tmpRange
        ForRange.Value = findWord(cnt)
        cnt = cnt + 1
    Next ForRange
    
    '検索結果があったらクリア
        '行の終わりを選択
        'すべての終わり行を調査する　一番大きいものの行が最終
        rect2.EndRow = 9
        cnt = cateCnt - 2
        For Num = 1 To cnt
            If Cells(65535, Num).End(xlUp).Row > rect2.EndRow Then
                rect2.EndRow = Cells(65535, Num).End(xlUp).Row
            End If
        Next
        Range(Cells(rowcnt - 2, 1), Cells(rect2.EndRow, cateCnt + 1)).Delete

    '構造体宣言
    Dim actData() As actStruct
    ReDim actData(cateCnt + 1)
    '検索、項目ワードを受け渡し用構造体へ
    For i = 1 To cateCnt
        actData(i).cate = cateWord(i)
        actData(i).find = findWord(i)
    Next i
    actData(0).Ccnt = cateCnt
    'シート名前用配列を確保　設定の5列目がシート名
    sCnt = Worksheets("設定").Cells(65535, 5).End(xlUp).Row
    If sCnt > cateCnt Then
        ReDim Preserve actData(sCnt)
    End If
    'シートネーム読み込み　配列へ
    With Worksheets("設定")
        Set tmpRange = .Range(.Cells(1, 5), .Cells(.Cells(65535, 5).End(xlUp).Row, 5))
        cnt = 1
        For Each ForRange In tmpRange
            actData(cnt).SheetName = ForRange.Value
            cnt = cnt + 1
        Next ForRange
    End With

    '検索　　////////////////////////////////////////////////////////////////////
    cnt = 1
    For cnt = 1 To sCnt
        actData(0).rcnt = rowcnt
        actData(0).SheetName = actData(cnt).SheetName
        rowcnt = actFind(actData())
    Next cnt

    
    
    Worksheets("検索").Activate
    Cells(1, 1).Select
    
End Function

'actFind
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
Function actFind(actData() As actStruct) As Integer
    Dim Num As Integer
    Dim rowcnt As Integer
    Dim fRect As SellRect
    Dim tmpRange As Range
    Dim Flag As Integer
    Dim obj As Object
    Dim sCnt As Integer
    '行数引継ぎ
    rowcnt = actData(0).rcnt
    With Worksheets(actData(0).SheetName)
        .Activate
        'カテゴリ行の最初から最後まで選択　　///////////////検索する座標を記憶
        fRect.cName = actData(4).cate  ' 1 -> No
        fRect = find_set(fRect)

        '検索範囲を選択
        Set obj = .Range(.Cells(fRect.stRow, 1), .Cells(fRect.EndRow, 1))
        '///////////////////////////////　検索する座標を記憶　END
        '///////////////////////////////　検索メイン
        For Each ForRange In obj
            Flag = 0
            For Num = 0 To actData(0).Ccnt
                '検索ワードが　空ならば無視
                If Not (actData(Num + 1).find = "") Then
                    '一致したのがあれば１　一致しなければ０　大文字・小文字を区別しない
                    If InStr(1, ForRange.Offset(0, Num).Value, actData(Num + 1).find, vbTextCompare) > 0 Then
                        Flag = 1
                    Else
                        Flag = 0
                        Exit For
                    End If
                End If
            Next Num
            'flag = 1 ならペースト
            If Flag = 1 Then
                ForRange.EntireRow.Copy
                With Worksheets("検索")
                    .Rows(rowcnt).PasteSpecial
                    rowcnt = rowcnt + 1
                End With
                
            End If
        Next ForRange
    End With
    actFind = rowcnt
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
