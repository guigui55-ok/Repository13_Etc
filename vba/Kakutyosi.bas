Attribute VB_Name = "Kakutyosi"
Sub 拡張子削除Down()
    Call Kakutyosi_Erase
End Sub


'メイン
'使い方：EraseString,SheetName,を入力して、開始セルを手動(マウス)で選択する。
'その後VBEのFunction Kakutyosi_Erase(以下関数内で)[F5]キーを押すと実行される
'
Public Function Kakutyosi_Erase()
    Dim RangeTemp As Range
    Dim RangeBegin As Range
    Dim EraseString As String
    Dim SheetName As String
    
    EraseString = ".mp4"
    SheetName = "D_AV"
    Set RangeBegin = Selection
    Set RangeTemp = Kakutyosi_Erase_SetRange(SheetName)
    '消す
    Call Kakutyosi_Erase_Core(RangeBegin, RangeTemp, EraseString)
    
    Set RangeBegin = Nothing
    Set RangeTemp = Nothing
    MsgBox "ok"
End Function

Function Kakutyosi_Erase_SetRange(SheetName As String) As Range
    Dim RowNow, RowEnd As Long
    Dim ColNow, ColEnd As Long
    Dim RowTemp As Long
    Dim ColTemp As Long
    
    'レンジセット
    RowNow = Selection.Row
    ColNow = Selection.Column
    RowEnd = Sheets(SheetName).Cells(RowNow, ColNow).End(xlDown).Row
    ColEnd = ColNow
    If RowNow = Rows.Count Then
        MsgBox "End Row = " & RowEnd
        Stop
    End If

    Set Kakutyosi_Erase_SetRange = Range( _
             Sheets(SheetName).Cells(RowNow, ColNow).Address, _
             Sheets(SheetName).Cells(RowEnd, ColEnd).Address _
    )
End Function

Function Kakutyosi_Erase_Core( _
    BeginRange As Range, EraseRange As Range, EraseString As String) As Boolean
    Dim ForRange As Range
    Dim cnt As Integer
    
    cnt = 0
'    Debug.Print EraseRange.Address
    For Each ForRange In EraseRange
        If InStr(1, ForRange.Value, EraseString, vbBinaryCompare) > 0 Then
            cnt = cnt - 1
            ForRange.Value = ""
        Else
'            Debug.Print ForRange.Offset(Cnt, 0).Address
'            Debug.Print ForRange.Address
'            Debug.Print ForRange.Offset(0, 0).Value
            ForRange.Offset(cnt, 0).Value = ForRange.Offset(0, 0).Value
            If Not cnt = 0 Then
                ForRange.Offset(0, 0).Value = ""
            End If
        End If
        If (ForRange.Value = "") And (ForRange.Offset(1, 0).Value = "") Then
            Exit For
        End If
    Next
    Set ForRange = Nothing
    cnt = 0
End Function


