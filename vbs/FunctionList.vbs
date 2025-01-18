'Start Function

'/////////////////////////////////////////////////////////////////////////
'デスクトップパス
Function GetDeskTopPath()
Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
GetDeskTopPath = objWShell.SpecialFolders("Desktop")
Set objWshell = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル削除
Function DeleteFileF(FullPath) 'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	If ExistsFile(FullPath) Then
		objFso.DeleteFile FullPath , True
		DeleteFileF = True
	Else
		DeleteFileF = False
	End If
Else
	DeleteFileF = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'フォルダ削除
Function DeleteFolderF(FullPath) 'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	If ExistsFolder(FullPath) Then
		objFso.DeleteFolder FullPath , True
	Else
		DeleteFolderF = False
	End If
Else
	DeleteFolderF = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイルの移動
Function MoveFileF(BaseFullPath,MoveFullPath) 'As Boolean
Dim objFso , tFlag
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	If ExistsFile(BaseFullPath) Then
		If ExistsFile(MoveFullPath) Then
			'移動先にすでに存在する
			objFso.DeleteFile MoveFullPath, True
			If Err.Number = 0 Then
				'削除された
				objFso.CopyFile BaseFullPath, MoveFullPath
				objFso.DeleteFile BaseFullPath, True
				MoveFileF = True
			Else
				'削除されず
				MoveFileF = False
			End If
		Else
			objFso.CopyFile BaseFullPath, MoveFullPath 
			objFso.DeleteFile BaseFullPath, True
			MoveFileF = True
		End IF
	Else
		'ファイル、フォルダが存在しない
		MoveFileF = False
	End If
Else
	MoveFileF = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'フォルダの移動 （移動先に存在したら、先削除->先へコピー->元削除 、存在しなければ 先へコピー->元削除）
Function MoveFolderF(BaseFullPath,MoveFullPath) 'As Boolean
Dim objFso , tFlag , objFso2
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	If ExistsFolder(BaseFullPath) Then
		If ExistsFolder(MoveFullPath) Then
			'移動先にすでに存在する
			objFso.DeleteFolder MoveFullPath, True
			If Err.Number = 0 Then
				'削除された
				objFso.CopyFolder BaseFullPath, MoveFullPath
				objFso.DeleteFolder BaseFullPath, True
				MoveFolderF = True
			Else
				'削除されず
				MoveFolderF = False
			End If
		Else
			objFso.CopyFolder BaseFullPath, MoveFullPath 
			objFso.DeleteFolder BaseFullPath, True
			MoveFolderF = True
		End IF
	Else
		'ファイル、フォルダが存在しない
		MoveFolderF = False
	End If
Else
	MoveFolderF = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイルリネーム 
Function RenameFile(Path , OldName , NewName) 'As Boolean
Dim objFso , FullPathOld , FullPathNew ,objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	FullPathOld = Path & "\" & OldName
	If ExistsFile(FullPathOld) Then
		FullPathNew = Path & "\" & NewName
		If ExistsFile(FullPathNew) Then
			'変更先ファイルが既に存在する
			'MsgBox FullPathNew & " = True"
			RenameFile = False
		Else
			'変更先ファイルが存在しない 場合リネーム
			Set objFile = objFso.GetFile(FullPathOld)
			'MsgBox "objFile.Name = " & objFile.Name
			objFile.Name = NewName
			RenameFile = True
		End If
	Else
		'変更元ファイルが存在しない
		RenameFile = False
	End If
End If
Set objFso = Nothing
End Function
'///////////////////////////////////////////////////////////////////////// 
'エラー出力
Function OutError(en , es) 'As 
	Dim FileName , Path , Flag
	Tstr = "Err.Number = " & en & vbNewLine & "Err.Description = " & es
	FileName = "ErrorLog.txt"
	Path = GetCurrentFolder & "\" & Filename
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
Set objFso = Nothing
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
Set objFso = Nothing
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
Set objWshell = Nothing
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
Set objWshell = Nothing
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
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'フォルダコピー
Function CopyFolder(CopyPath,PastePath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
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
End If
Set objFso = Nothing
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
		'MsgBox FullPath  & " が存在していません。"
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
For Each OSInfo In OSInfoCollection 'ループはコレクションの1番目を参照するためのみの役割
'    WScript.Echo "オペレーティングシステム：" & OSInfo.Caption
'    WScript.Echo "バージョン：" & OSInfo.Version
'    WScript.Echo "サービスパック：" & OSInfo.CSDVersion
'    WScript.Echo "システムディレクトリ：" & OSInfo.SystemDirectory
'    WScript.Echo "システムドライブ：" & OSInfo.SystemDrive
'    WScript.Echo "仮想メモリ容量：" & OSInfo.TotalVirtualMemorySize & "Bytes"
'    WScript.Echo "物理メモリ容量：" & OSInfo.TotalVisibleMemorySize & "Bytes"
	GetOSName = OSInfo.Caption
Next
End Function
'/////////////////////////////////////////////////////////////////////////
'GetOSName の名前を簡略化 Select Case用
Function OSNameShort(OSName)	'As String
Dim List,tV
List = Array("XP","Vista","7")
For Each tV In List
	If Instr(1,OSName,tV,VbBinaryCompare)>0 Then
		OSNameShort = tV
	End If
Next
End Function
'/////////////////////////////////////////////////////////////////////////
'ショートカットを作成
Function MakeShortCut(BasePath , FileName , MakePath , ShortCutTitle,IconFilePath,IconNum)	'As Boolean
Dim objWshell , objShortcut ,BaseFullPath
Set objWshell = WScript.CreateObject("WScript.Shell")
BaseFullPath = BasePath & "\" & FileName
'BasePath フォルダ存在チェック
'BaseFullPath ファイル存在チェック
'MakePath フォルダ存在チェック
If Err.Number = 0 Then
    'strDesktopPath = objWshell.SpecialFolders("Desktop")  'デスクトップに　ショートカット先
    'strWindowsPath = objWshell.ExpandEnvironmentStrings("%WINDIR%") 'メモ帳の場所 ショートカット元の場所
    Set objShortcut = objWshell.CreateShortcut(MakePath & "\" & ShortCutTitle & ".lnk") 'ショートカット作成先
    objShortcut.Description = ShortCutTitle
'    objShortcut.HotKey = "CTRL+ALT+N"
    objShortcut.IconLocation = IconFilePath & "," & IconNum
    objShortcut.TargetPath = BasePath & "\" & FileName	'ショートカット作成元
    objShortcut.WorkingDirectory = MakePath
    objShortcut.Save
	MakeShortCut = True
Else
	MakeShortCut = False
End If
Set objFso = Nothing
Set objShortcut = Nothing
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
Set objFso = Nothing
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
Function GetFunctionList(FullPath)	'As String
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


'/////////////////////////////////////////////////////////////////////////
'ファイル存在チェック
Function ExistsFileL(FullPath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(FullPath) = True Then
        '存在している。
		ExistsFileL = True
    Else
        '存在していない
		ExistsFileL = False
		'MsgBox FullPath  & " が存在していません。"
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル書き込み　なければ作る
Function WriteMakeTextL(FullPath,Wstr) 'As Boolean
Dim objFso , objStm , Rstr , tFlag
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFileL(FullPath) Then
	'存在する場合読み込んで書き込む
	Rstr = ReadFileL(FullPath)
	Rstr = Rstr & Vbnewline & Wstr
	tFlag = WriteFileL(FullPath,Rstr)
Else
	'存在しない場合は作って書き込み
	Set objStm = objFso.CreateTextFile(FullPath, true)
		objStm.Close()
	If Err.Number = 0 Then
		tFlag = WriteFileL(FullPath,Wstr)
	End If
End If
Set objFso = Nothing
Set objStm = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル書き込み
Function WriteFileL (FullPath,Wstr)	'As Boolean
Dim objFso, objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
'MsgBox Wstr
If ExistsFileL(FullPath) Then
    If Err.Number = 0 Then
		Set objFile = objFSO.OpenTextFile(FullPath, 2, True)
        objFile.WriteLine(Wstr)
		objFile.Close
		WriteFileL = True
	Else
		WriteFileL = False
	End If
Else
	WriteFileL = False
End If
Set objFile = Nothing
Set objFSO = Nothing

End Function
'/////////////////////////////////////////////////////////////////////////
'ファイル読み込み
Function ReadFileL (FileFullPath)	'As Boolean
Dim objFso, objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFileL(FileFullPath) Then
    If Err.Number = 0 Then
		Set objFile = objFso.OpenTextFile(FileFullPath)
		ReadFileL = objFile.ReadAll
		objFile.Close
	Else
		ReadFileL = False
	End If
Else
	ReadFileL = False
End If
Set objFile = Nothing
Set objFSO = Nothing
End Function
