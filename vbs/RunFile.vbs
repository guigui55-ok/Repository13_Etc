'Sub test_OnClick  HTML
Sub test()
'Msgbox "ボタンを押してください。",5+64+256+4096,"VBScriptテスト"
	MsgBox GetCurrentDirectory

End Sub

'MsgBox "RunFile.vbs Read."
Function ScriptFullPATH() 'CurrentDirectory　とおなじ
	Dim strPATH
Dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")
	StrPATH = objFso.GetAbsolutePathName("")
	'ScriptFullPATH = Left(strPATH,InStrRev(strPATH,"\")-1)
	ScriptFullPATH = StrPATH
End Function




'EXEファイルを実行する
Function RunExe(Name)
	Dim WSHShell,FullPath
		Set WSHShell = CreateObject ("Wscript.Shell")
	MsgBox ScriptFullPATH
	FullPath = GetCurrentDirectory & "\" & Name
	If ExistsFile(FullPath) Then
		WSHShell.Run "C:\Windows\notepad.exe", 3, True
	Else
		MsgBox FullPath & " が存在しません。" & vbnewline & "終了します。"
	End If
End Function

'EXEファイルにパラメータを指定して実行する
Function RunExeWithParam(FileName)
	MsgBox GetCurrentDirectory
	Dim WSHShell
	Set WSHShell = CreateObject ("Wscript.Shell")
	WSHShell.Run "C:\Windows\notepad.exe '" & FileName & "'", 3, True
End Function

Function GetCurrentDirectory()
	Dim WSHShell
	Set WSHShell = CreateObject ("Wscript.Shell")
	GetCurrentDirectory = WSHShell.CurrentDirectory
End Function

'ファイル存在チェック
Function ExistsFile (FullPath)	'As Boolean
Dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")
'Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(FullPath) = True Then
        '存在している。
		ExistsFile = True
    Else
        '存在していない
		ExistsFile = False
		'MsgBox FullPath  & " が存在していません。"
    End If
Set objFso = Nothing
End Function