
'他のスクリプトの関数を実行する
'関数を呼び出す側「a.vbs」
 
Option Explicit
Dim ReadDir
ReadDir = GetCurrentDirectory()
Dim ReadFileName
ReadFileName = "FunctionList_sjis.vbs"
Dim ReadPath
ReadPath = ReadDir + "\" + ReadFileName
'MsgBox ReadDir
Include(ReadPath)	'vbsの読み込み
Include(ReadFileName)	'vbsの読み込み

Dim WriteData
WriteData = GetOSName() 'vbsの関数を実行
MsgBox WriteData
Dim WritePath
WritePath = GetCurrentDirectory()
WritePath = WritePath + "\" + "Log.txt"
Dim Ret
Ret = WriteFile(WritePath,WriteData)
Dim Msg
Msg = "WriteFile : " & CStr(Ret) & " ,  Path = " & WritePath
MsgBox Msg

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
