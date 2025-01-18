Option Explicit
dim gLogger
 ' ==============================================
'他のスクリプトの関数を実行する
Dim ReadDir :ReadDir = GetCurrentDirectory()
Dim ReadFileName :ReadFileName = "FunctionList_sjis.vbs"
Dim ReadPath: ReadPath = ReadDir + "\" + ReadFileName
'MsgBox ReadDir
Include(ReadPath)	'vbsの読み込み
Include(ReadFileName)	'vbsの読み込み
' ==============================================
'ログ書き込み用のクラス
Dim objLogger
Set objLogger = New Logger
set gLogger = objLogger
objLogger.SetPath(ReadDir + "\MoveFilesJpg_Log.txt")
'Call objLogger.AddLog("AddLog")
' ==============================================
'コピー元のディレクトリを設定する
Dim FileDir :FileDir = "C:\Users\OK\source\repos\Repository5_Etc\vbs\test_base"
'ディレクトリのファイルリストを取得する
Dim FileArray :FileArray = GetListFile(FileDir)
'配列を文字列にする
'Dim Buf :Buf = AryJoin(FileArray,",")
'MsgBox Buf
'コピー先のディレクトリを設定する
Dim CopyDir: CopyDir = "C:\Users\OK\source\repos\Repository5_Etc\vbs\test2\jpg"
'対象の拡張子
Dim TargetExt:TargetExt = "jpg"

Dim FileName
Dim BaseFilePath
Dim CopyDistPath:CopyDistPath = CopyDir
Dim Flag
Dim i
Dim RetStr
For i = 0 To UBound(FileArray)
	FileName = FileArray(i)
	BaseFilePath = FileDir + "\" + FileName
	CopyDistPath = CopyDir + "\" + FileName
	If ExtensionIsMatch(BaseFilePath, TargetExt) Then
		Flag = CopyFileWithRename(BaseFilePath,CopyDistPath)
	End If
	'Exit For
Next
'RetStr = CStr(Flag)
'RetStr = RetStr + " , " + BaseFilePath + "  =>  " + CopyCopyDistPathPath
'MsgBox RetStr
'Dim WriteData :WriteData = ""
'WriteData = RetStr
'WriteData = GetOSName() 'vbsの関数を実行
'MsgBox WriteData
' ==============================================
'ログファイルに書き込む
Dim WritePath :WritePath = GetCurrentDirectory()
'日付を取得する
Dim DateStr :DateStr = GetDateStr()
WritePath = WritePath + "\_" + DateStr + "_Log.txt"
'ファイルを書き込む
'Dim Ret :Ret = WriteFile(WritePath,WriteData)
'Dim Ret
'Dim Msg
'Msg = "WriteFile : " & CStr(Ret) & " ,  Path = " & WritePath
'MsgBox Msg
' ==============================================
'/////////////////////////////////////////////////////////////////////////
Function CopyFileWithRename(CopyFullPath , PasteFullPath)	'As Boolean
	Dim RetFlag
	Dim DistPath
	If ExistsFile(PasteFullPath) Then
		'コピー先にコピー元と同名のファイルが存在する場合は、
		'ファイル名の後ろに「(番号)」を付与してコピーする、
		'同じ番号がある場合はカウントアップする
		Dim NewPasteFullPath:NewPasteFullPath = GetCopyNextFileNameIfCopyFileFormat(PasteFullPath,0)
		DistPath = NewPasteFullPath
	Else
		'コピー先にコピー元と同名のファイルがない場合はそのままコピーする
		DistPath = PasteFullPath
	End IF
	'MsgBox DistPath
	RetFlag = CopyFile(CopyFullPath,DistPath)
	CopyFileWithRename = RetFlag
	
	Dim LogStr:LogStr = "Copy: " + CStr(RetFlag) + "," + BaseFilePath + " => " + DistPath
	AddLog(LogStr)
	'MsgBox LogStr
	DPErr("CopyFileWithRename")
End Function
'/////////////////////////////////////////////////////////////////////////
' C:\DirName\filename(2).ext -> filename(3).ext にする
Function GetCopyNextFileNameIfCopyFileFormat(FilePath,Num)
	'存在しないときはそのまま返す
	If Not ExistsFile(FilePath) Then
		GetCopyNextFileNameIfCopyFileFormat = FilePath
		Exit Function
	End If
	Dim DirPath:DirPath = GetDirectoryFromPath(FilePath)
	Dim FileBaseNameNotExt:FileBaseNameNotExt = GetBaseName(FilePath)
	Dim Ext:Ext = GetExt(FilePath)
	Dim AfterStr:AfterStr = ")"
	Dim BeforeStr:BeforeStr = " ("
	Dim CountUpFileName:CountUpFileName = CountUpNumber(FileBaseNameNotExt,BeforeStr,AfterStr)
	Dim NewFilePath:NewFilePath = DirPath + CountUpFileName + "." + Ext
	If Num > 10000 Then
		'無限ループ防止用に上限を設ける
		GetCopyNextFileNameIfCopyFileFormat = ""
		Exit Function
	End If
	'次の番号が存在して、上書きしないように、再帰的に実行する
	Num = Num + 1
	GetCopyNextFileNameIfCopyFileFormat = GetCopyNextFileNameIfCopyFileFormat(NewFilePath,Num)
	'MsgBox GetCopyNextFileNameIfCopyFileFormat
End Function
'/////////////////////////////////////////////////////////////////////////
'BeforeStr と AfterStr の間が数字なら、カウントアップして返す
' filename(2) -> filename(3) にする
Function CountUpNumber(FileBaseNameNotExt,BeforeStr,AfterStr)
	Dim LastChar: LastChar = Right(FileBaseNameNotExt, Len(AfterStr))
	Dim Num
	Dim NewBaseName
	If StrComp(LastChar,AfterStr) = 0 Then
		'最後から1文字目までを取得する filename(2
		Dim RightPos : RightPos = Len(FileBaseNameNotExt)-Len(AfterStr)
		Dim Buf : Buf = Left(FileBaseNameNotExt, RightPos)
		'BeforeStr までの位置を取得
    	Dim LeftPos :LeftPos = InStrRev(FileBaseNameNotExt, BeforeStr)
		'数字を取得する filename(2 => 2
		Dim NumStr: NumStr = Mid(FileBaseNameNotExt,LeftPos + Len(BeforeStr), RightPos - Len(AfterStr) - LeftPos)
		'数値であればカウントアップし、そうでなければ1を付与する	
		If IsNumeric(NumStr) Then
			Num = CInt(NumStr) + 1
		Else
			Num = 1
		End If
		'BeforeStr までの文字列を BaseName とする
		NewBaseName = Left(FileBaseNameNotExt,LeftPos - 1)
		'MsgBox NewBaseName
	Else
		Num = 1
		NewBaseName = FileBaseNameNotExt
	End If
	Dim AddStr:AddStr = BeforeStr + CStr(Num) + AfterStr
	CountUpNumber = NewBaseName + AddStr
	'MsgBox CountUpNumber
End Function
'/////////////////////////////////////////////////////////////////////////
'拡張子なしのファイル名を取得する
Function GetBaseName(Path)
	Dim objFileSys
	'ファイルシステムを扱うオブジェクトを作成
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	GetBaseName = objFileSys.getBaseName(Path)
	Set objFileSys = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
Function ExtensionIsMatch(Path,Ext)'As Boolean
	Dim PathExt:PathExt = GetExt(Path)
	'Dim Val:Val = PathExt & " , " & Ext
	'AddLog(Val)
	'MsgBox Val
	If PathExt = Ext Then
		ExtensionIsMatch = True
	Else
		ExtensionIsMatch = False
	End If
End Function
'/////////////////////////////////////////////////////////////////////////
'拡張子を取得する
Function GetExt(Path)
	Dim objFileSys
	'ファイルシステムを扱うオブジェクトを作成
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	GetExt = objFileSys.GetExtensionName(Path)
	Set objFileSys = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'ディレクトリ名を取得する
Function GetDirectoryFromPath(Path)
    Dim DirName, pos
    pos = InStrRev(Path, "\")
    DirName = Left(Path, pos)
    'Dim FileName:FileName = Mid(Path, pos + 1)
    GetDirectoryFromPath = DirName
End Function
'/////////////////////////////////////////////////////////////////////////
 Class Logger
    Public gDir
    Public gFileName
	Public gFilePath
	Private pClassName
    Private Sub Class_Initialize
        pClassName = "Logger."
		gDir = ""
		gFileName = ""
		gFilePath = ""
    End Sub
	public Sub SetPath(Path)
		gFilePath = Path
	End Sub
	Public Sub AddLog(Value)
		'日付を取得する
		'Dim DateStr :DateStr = GetDateStr()
		Dim DateStr :DateStr = CStr(Now())
		Value = DateStr + " " + Value
		Dim Ret :Ret = WriteFile(gFilePath,Value)
		DPErr(pClassName & "AddLog")
	End Sub
 End Class
'/////////////////////////////////////////////////////////////////////////
Function AddLog(Value)
	gLogger.AddLog(Value)
End Function
'/////////////////////////////////////////////////////////////////////////
Function Include(strFile)
	'strFile：読み込むvbsファイルパス 
	Dim objFso, objWsh, strPath
	Set objFso = Wscript.CreateObject("Scripting.FileSystemObject")
	
	'外部ファイルの読み込み
	Set objWsh = objFso.OpenTextFile(strFile)
	ExecuteGlobal objWsh.ReadAll()
	'objWsh.ReadAll()
	objWsh.Close
 
	Set objWsh = Nothing
	Set objFso = Nothing
	DPErr("Include")
End Function
'/////////////////////////////////////////////////////////////////////////
'カレントディレクトリを取得する
Function GetCurrentDirectory()
	Dim objShell
	Dim curDir
	Set objShell = CreateObject( "WScript.Shell" )
	curDir = objShell.CurrentDirectory

	Set objShell = Nothing
	GetCurrentDirectory = curDir
	DPErr("GetCurrentDirectory")
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
