'Start Function
'///////////////////////////////////////////////////////////////////////// 
Function OutError(en , es) 'As 
	Dim FileName , Path , Flag
	Tstr = "Err.Number = " & en & vbNewLine & "Err.Description = " & es
	FileName = "Error.txt"
	Path = GetCurrentFolder & Filename
	Flag = WriteFile(Path,Tstr)
End Function
'///////////////////////////////////////////////////////////////////////// 
'フォルダリスト取得
Function GetListSubFolder(FullPath)	'As String()
Dim objFso , Cnt , List() , fV , objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFolder(FullPath) Then
	Set objFile = objFso.GetFolder(FullPath)
	Cnt = 0
	For Each fV In objFile.SubFolders
		Redim Preserve List(Cnt)
		List(Cnt) = fV.Name
		Cnt = Cnt + 1
	Next
	GetListSubFolder = List
Else
	Redim List(0)
	GetListSubFolder = List
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイルリスト取得
Function GetListFile(FullPath)	'As String()
Dim objFso , Cnt , List() , fV , objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFolder(FullPath) Then
	Set objFile = objFso.GetFolder(FullPath)
	Cnt = 0
	For Each fV In objFile.Files
		Redim Preserve List(Cnt)
		List(Cnt) = fV.Name
		Cnt = Cnt + 1
	Next
	GetListFile = List
Else
	Redim List(0)
	GetListFile = List
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'配列を文字列に　VbNewLineで区切る
Function AryComp(Tstra)		'As String
Dim Tstr
If AryCheckZero(Tstra) Then
	Tstr = ""
	For i = 0 to Ubound(Tstra)
		Tstr = Tstr & Tstra(i) & VbNewLine
	Next
	AryComp =  Tstr
Else
	AryComp = ""
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'配列を出力
Function OutAry(Tstra)	'AS
Dim Tstr
If AryCheckZero(Tstra) Then
	Tstr = ""
	For i = 0 to Ubound(Tstra)
		Tstr = Tstr & Tstra(i)
	Next
	MsgBox Tstr
Else
	MsgBox "AryCheckZero = 0"
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'配列があるかチェック
Function AryCheckZero(Tstra)	'As Boolean
On Error Resume Next
For i = 0 to UBound(Tstra)

Next
If Err.Number = 0 Then
	AryCheckZero = True
Else
	AryCheckZero = False
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'文字列を入力
Function FormInput(Title)	'As 
FormInput = InputBox(Title)
'MsgBox (Input & "を入力しました。")
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル実行
Function RunFile (FullPath)	'As
'On Error Resume Next
Dim objWshell ,Z
If Err.Number = 0 Then
	If ExistsFile(FullPath) Then
		Set objWshell = WScript.CreateObject("WScript.Shell")
		'ファイル存在チェック
		objWshell.Run FullPath, vbNormalFocus,True
	Else
		MsgBox "File Not Foune." & vbnewline & FullPath
	End If
Else
End if
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル実行
Function RunFile2(FullPath)	'As
Dim objWshell ,Z
		Set objWshell = WScript.CreateObject("WScript.Shell")
		Set oExec = objWshell.Exec(FullPath)

		objWshell.Run "test3.vbs", vbNormalFocus, True
		Wscript.Echo ExistsFile(FullPath) & vbnewline &  FullPath
		objWshell.Run FullPath, vbNormalFocus, True
'MsgBox FullPath
End Function
'/////////////////////////////////////////////////////////////////////////
'フォルダ存在チェック
Function ExistsFolder(Path)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If objFso.FolderExists(Path) = True Then
    '存在している。
	ExistsFolder = True
Else
    '存在していない
	ExistsFolder = False
	'MsgBox Path  & " が存在していません。"
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'フォルダコピー
Function CopyFolder(CopyPath,PastePath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFolder(CopyPath) = True Then
	If ExistsFolder(PastePath) <> True Then
		' コピー先フォルダが存在しないときは作成する
		objFSO.CreateFolder(PastePath)
	Else
		CopyFolder = False
	End IF
	'WScript.echo Copypath & vbnewline & PastePath
	' フォルダコピー
	    objFso.CopyFolder CopyPath, PastePath
		CopyFolder = True
Else
	CopyFolder = False
	'MsgBox CopyPath  & " が存在していません。"
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル存在チェック
Function ExistsFile (FullPath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(FullPath) = True Then
        '存在している。
		ExistsFile = True
    Else
        '存在していない
		ExistsFile = False
		'MsgBox FullPath  & " が存在していません。" ''終了していない文字列型の定数です。
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイルコピー
Function CopyFile (CopyFullPath , PasteFullPath)	'As Boolean
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
Function GetCurrentFolder()	'As String
Dim objWshell
Set objWshell = WScript.CreateObject("WScript.Shell")
	GetCurrentFolder = objWshell.CurrentDirectory 
Set objWshell = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'OSの種類の名前を取得
Function GetOSName()	'As String
Dim OSInfoCollection
Set OSInfoCollection = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
dim RetStr
RetStr = ""
For Each OSInfo In OSInfoCollection 'ループはコレクションの1番目を参照するためのみの役割
	RetStr = RetStr + "オペレーティングシステム：" & OSInfo.Caption + vbNewLine
'    WScript.Echo "オペレーティングシステム：" & OSInfo.Caption
	RetStr = RetStr + "バージョン：" & OSInfo.Version + vbNewLine
'    WScript.Echo "バージョン：" & OSInfo.Version
	RetStr = RetStr + "サービスパック：" & OSInfo.CSDVersion + vbNewLine
'    WScript.Echo "サービスパック：" & OSInfo.CSDVersion
	RetStr = RetStr + "システムディレクトリ：" & OSInfo.SystemDirectory + vbNewLine
'    WScript.Echo "システムディレクトリ：" & OSInfo.SystemDirectory
	RetStr = RetStr + "システムドライブ：" & OSInfo.SystemDrive + vbNewLine
'    WScript.Echo "システムドライブ：" & OSInfo.SystemDrive
	RetStr = RetStr + "仮想メモリ容量：" & OSInfo.TotalVirtualMemorySize & "Bytes" + vbNewLine
'    WScript.Echo "仮想メモリ容量：" & OSInfo.TotalVirtualMemorySize & "Bytes"
	RetStr = RetStr + "物理メモリ容量：" & OSInfo.TotalVisibleMemorySize & "Bytes" + vbNewLine
'    WScript.Echo "物理メモリ容量：" & OSInfo.TotalVisibleMemorySize & "Bytes"
	GetOSName = OSInfo.Caption
Next
GetOSName = RetStr
End Function
'/////////////////////////////////////////////////////////////////////////
'ショートカットを作成
Function MakeShortCut(BasePath , FileName , MakePath , ShortCutTitle)	'As Boolean
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
Function ReadFile (FileFullPath)	'As Boolean
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
Function WriteFile (FullPath,Wstr)	'As Boolean
Dim objFso, objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
'MsgBox Wstr
If ExistsFile(FullPath) Then
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
'関数リスト作成
Function GetFunctionList(FullPath)'As String
	Dim Title , FileName ,FileFullPath , ReadStr ,Aftstr
	'Title = "読み込みファイル名を入力（フルパスで）"
	Title = "******"
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
Function GetFunctionList_MainProcess(Rstr)	'As String
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
'/////////////////////////////////////////////////////////////////////////
'動作テスト
Function Testmsg() 'As
	MsgBox "Test"
End Function
