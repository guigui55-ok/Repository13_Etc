'Option Explicit
'On Error Resume Next

Dim  Flag , Tstr , Z
Dim FileName, FullPath


DirPath = GetCurrentDirectory()
FileName = "FunctionList.vbs"
FullPath = DirPath & "\" & FileName '入力ファイル名
FileName = GetCurrentFolder & "\" & "MakeFunctionList.txt"   '出力ファイル名
Tstr = GetFunctionList(FullPath) '読み込み
Flag = WriteFile(FileName,Tstr)

'FullPath = "D:\zzz\HowTo\Software\Test.vbs"	'入力ファイル名
Z = OutTestToFunctionList(FullPath)
Result = CStr(Flag)  & " : " & FullPath
MsgBox Result
'/////////////////////////////////////////////////////////////////////////
Function GetListFile(FullPath)
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject").GetFolder(FullPath).Files


End Function
'/////////////////////////////////////////////////////////////////////////
Function GetCurrentDirectory()
	Dim objShell
	Dim curDir

	' カレントディレクトリ取得.
	Set objShell = CreateObject( "WScript.Shell" )
	curDir = objShell.CurrentDirectory

	Set objShell = Nothing
	GetCurrentDirectory = curDir
End Function
'/////////////////////////////////////////////////////////////////////////
'文字列を入力
Function FormInput(Title)
FormInput = InputBox(Title)
'MsgBox (Input & "を入力しました。")
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル実行
Function RunFile (FullPath)
Dim objWshell
Set objWshell = WScript.CreateObject("WScript.Shell")
'ファイル存在チェック
objWshell.Run  FullPath
End Function
'/////////////////////////////////////////////////////////////////////////
'フォルダ存在チェック
Function ExistsFolder(Path)
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFso.FolderExists(Path) = True Then
        '存在している。
		ExistsFolder = True
    Else
        '存在していない
		ExistsFolder = False
    End If
End Function
'/////////////////////////////////////////////////////////////////////////
'フォルダコピー
Function CopyFolder(CopyPath,PastePath)
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFolder(CopyPath) = True Then
	If ExistsFolder(PastePath) <> True Then
		' コピー先フォルダが存在しないときは作成する
		objFSO.CreateFolder(PastePath)
	Else
		CopyFolder = False
	End IF
	WScript.echo Copypath & vbnewline & PastePath
	' フォルダコピー
	    objFso.CopyFolder CopyPath, PastePath
		CopyFolder = True
Else
	CopyFolder = False
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル存在チェック
Function ExistsFile (FullPath)
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(FullPath) = True Then
        '存在している。
		ExistsFile = True
    Else
        '存在していない
		ExistsFile = False
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイルコピー
Function CopyFile (CopyFullPath , PasteFullPath)
Dim objFso
'Wscript.echo copyfullpath
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFile(CopyFullPath) Then
	 ' ファイルコピー
	objFso.CopyFile CopyFullPath, PasteFullPath, True
	CopyFile = True
	Else
		'存在しない
	CopyFile = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'カレントフォルダ取得
Function GetCurrentFolder()
Dim objWshell
Set objWshell = WScript.CreateObject("WScript.Shell")
	GetCurrentFolder = objWshell.CurrentDirectory 
Set objWshell = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ショートカットを作成
Function MakeShortCut(BasePath , FileName , MakePath , ShortCutTitle)
Dim objWshell , objShortcut ,BaseFullPath
Set objWshell = WScript.CreateObject("WScript.Shell")
BaseFullPath = BasePath & "\" & FileName
'BasePath フォルダ存在チェック
'BaseFullPath ファイル存在チェック
'MakePath フォルダ存在チェック
If Err.Number = 0 Then
    'strDesktopPath = objWshell.SpecialFolders("Desktop")  'デスクトップに　ショートカット先
    'strWindowsPath = objWshell.ExpandEnvironmentStrings("%WINDIR%") 'メモ帳の場所 ショートカット元の場所
    Set objShortcut = objWshell.CreateShortcut(MakePath & "\" & ShortCutTitle & ".lnk")
    objShortcut.Description = ShortCutTitle
    objShortcut.HotKey = "CTRL+ALT+N"
    objShortcut.IconLocation = FileName
    objShortcut.TargetPath = MakePath & "\" & FileName
    objShortcut.WorkingDirectory = MakePath
    objShortcut.Save
	MakeShortCut = True
Else
	MakeShortCut = False
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル読み込み
Function ReadFile (FileFullPath)
Dim objFso, objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFile(FileFullPath) Then
    If Err.Number = 0 Then
		Set objFile = objFso.OpenTextFile(FileFullPath)
		ReadFile = objFile.ReadAll
		objFile.Close
	Else
		ReadFile = False
	End If
Else
	ReadFile = False
End If
Set objFile = Nothing
Set objFSO = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル書き込み
Function WriteFile (FullPath,Wstr)
Dim objFso, objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
'MsgBox Wstr
If True Then
    If Err.Number = 0 Then
		Set objFile = objFSO.OpenTextFile(FullPath, 2, True)
        objFile.WriteLine(Wstr)
		objFile.Close
		WriteFile = True
	Else
		WriteFile = False
	End If
Else
	WriteFile = False
End If
Set objFile = Nothing
Set objFSO = Nothing

End Function
'/////////////////////////////////////////////////////////////////////////
'文字列を連結

'/////////////////////////////////////////////////////////////////////////
'テストファイルからインクルード用関数リスト作成
Function OutTestToFunctionList(FullPath)
 Dim Fstr , Tstr , n , Rstr , FileName , Flag

Rstr = ReadFile(FullPath)
Fstr = "'Start Function"
n = InStr(1,Rstr,Fstr,vbBinaryCompare) -1
If n > 0 Then
	Tstr = Right(Rstr,Len(Rstr) - n)
	OutTestToFunctionList = Tstr
	FileName = "FunctionList.vbs"
	FullPath = GetCurrentFolder & "\" & FileName
	Flag = WriteFile(FullPath,Tstr)
Else
	OutTestToFunctionList = ""
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'関数リスト作成
Function GetFunctionList(FullPath)
Dim Title , FileName ,FileFullPath , ReadStr ,Aftstr
Title = "読み込みファイル名を入力（フルパスで）"
'FullPath = "D:\zzz\Software\Test.vbs"
'FileFullPath = FormInput(Title)
ReadStr = ReadFile(FullPath)
'MsgBox Instr(1,ReadStr,vbNewLine , vbBinaryCompare )
Aftstr = GetFunctionList_MainProcess(ReadStr)
'MsgBox AftStr
GetFunctionList = AftStr
End Function
'/////////////////////////////////////////////////////////////////////////
'関数リスト作成メイン
Function GetFunctionList_MainProcess(Rstr)
Dim n , nbfo1, nbfo2 , Cnt , Memo() , Tstr , nbfo , naft , Kaigyo , nkaigyo , Max , tn , Cnt2 , Memo2()
'Kaigyo = vbNewLine
Kaigyo = vbCrLf
If Len(Rstr) > 0 Then
	Max = Len(Rstr)
	Tstr = ""
	n = Instr(1,Rstr, VbNewLine & "Function ",vbBinaryCompare)
	If n > 0 Then
	Cnt = 0
	Do While (n > 0)
		'あったらメモる
		Redim Preserve Memo(Cnt)
		Memo(Cnt) = n
		Cnt = Cnt + 1
		'改行＋Function をさがし、改行のみと一致したらメモする
		Cnt2 = 0
		nbfo1 = InStrRev( Rstr , Kaigyo & "'", n , vbBinaryCompare)
		nbfo2 = InStrRev( Rstr , Kaigyo , n , vbBinaryCompare)
		Do While (nbfo1 = nbfo2) And (nbfo1 > 0) And (nbfo2 > 0) 'あったらその前方方向（逆方向）に検索
			'メモがあったらメモる
			tn = nbfo1
			Redim Preserve Memo2(Cnt2)
			Memo2(Cnt2) = nbfo1
			Cnt2 = Cnt2 + 1
			'Memo(Cnt) = nbfo1
			nbfo1 = InStrRev( Rstr , Kaigyo & "'", tn , vbBinaryCompare)
			nbfo2 = InStrRev( Rstr , Kaigyo , tn , vbBinaryCompare)
		Loop
		If Cnt2 > 0 Then '順番おかしくなるので一時保存して追記
			For i = 0 to Ubound(Memo2)
				Redim Preserve Memo(Cnt)
				Memo(Cnt) = Memo2(i)
				Cnt = Cnt + 1
			Next
		End If
		n = n + 1
		If Max <= n Then Exit Do
		n = Instr(n,Rstr, Kaigyo & "Function ",vbBinaryCompare)
	Loop
		'メモしたものをもとにその行を文字列へ
		If Cnt > 0 Then
			For i = 0 to Ubound(Memo) 
				nbfo = Memo(i)
				naft = Instr(nbfo + 1, Rstr , Kaigyo , vbBinaryCompare)
				If (nbfo > 0) And (naft > 0) Then
					Tstr = Tstr & Mid(Rstr, nbfo , naft - nbfo)
				End If
			Next
		End If
		GetFunctionList_MainProcess = Tstr
	Else 'Function が一つも存在しない
		GetFunctionList_MainProcess = Tstr
	End If
End If
End Function