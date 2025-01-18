
CopyFIle("")

'*********************************************************
'用途: ファイルをコピーする
'受け取る値: なし
'戻り値: おみくじの文字列（String）
'*********************************************************
Function CopyFile(Path)
On Error Resume Next
    Dim strOmikujis, strOmikuji
    'おみくじの一覧を格納した配列
    'strOmikujis = Array("大吉")
    strOmikujis(2) = strOmikujis(2)
    'おみくじを引く（ランダムに1つ取り出す）
    strOmikuji = strOmikujis(Int((UBound(strOmikujis) + 1) * Rnd))
    '結果を戻り値として返す
    GetOmikuji = strOmikuji

DPErr()
End Function

'//////////////////////////////////////////////////////////////////////////
Function DPErr(FuncName)
    Dim msg
    If Err.Number <> 0 Then
        msg = "Error : " & Err.Number & " : " & Err.Description
        msg = msg + " , Function = " & FuncName
        'msg = msg + " , Source = " & Err.Source
        'msg = msg + " , Erl = " & CStr(Erl)
        msgbox msg
    End If
End Function
'//////////////////////////////////////////////////////////////////////////