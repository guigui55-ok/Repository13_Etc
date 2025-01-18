Attribute VB_Name = "Module5"
'find_set


'//////////////////////////////////////////////////////////////////////////////////////////////////////// find_set

Public Function find_set(ByRef FindRect As SellRect) As SellRect
    Dim findObj As Object
    
    '値が一致するセルを探す
    Set findObj = Cells.find(FindRect.cName, lookat:=xlWhole)
    If findObj Is Nothing Then
        'なければメッセージボックス表示　＆　フラグゼロにする
        MsgBox FindRect.cName & "Not Found"
        FindRect.Fflag = 0
        Exit Function
    Else
        'フラグが３ならば　検索値のセルの位置を記憶して　フラグを１にする
        If FindRect.Fflag = 3 Then
            FindRect.tmpRow = findObj.Row
            FindRect.tmpCol = findObj.Column
            FindRect.Fflag = 1
            find_set = FindRect
            Exit Function
        End If
        'フラグが１ならば　検索値のセルの位置とその最下行の位置を記憶する
        FindRect.Fflag = 1
        FindRect.stRow = findObj.Row
        FindRect.stCol = findObj.Column
        FindRect.EndRow = Cells(65535, FindRect.stCol).End(xlUp).Row
        FindRect.EndCol = Cells(65535, FindRect.stCol).End(xlUp).Column
        'Cells(FindRect.stRow, FindRect.stCol).Select
    End If
    'findObj.Select
    find_set = FindRect
End Function
'//////////////////////
