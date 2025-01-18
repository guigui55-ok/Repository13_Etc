Option Explicit
On Error Resume Next

Dim objWshShell     ' WshShell オブジェクト

Set objWshShell = WScript.CreateObject("WScript.Shell")
If Err.Number = 0 Then
    'WScript.Echo "現在のフォルダは " & objWshShell.CurrentDirectory & " です。"
Else
    WScript.Echo "エラー: " & Err.Description
End If

'===================
'move to
dim moveFolder
moveFolder = objWshShell.CurrentDirectory + "\jpg"
'msgbox "moveFolder : "+movefolder
'===================

dim fso
set fso = createObject("Scripting.FileSystemObject")

dim folder
set folder = fso.getFolder(objWshShell.CurrentDirectory)

' ファイル一覧
dim writePath
writePath = objWshShell.CurrentDirectory + "\0log.txt"
'msgbox writepath
dim list
dim file
list = ""

dim moveBase
dim moveTo
dim flag
flag = false
dim CheckPath
dim checkpath2
dim count
count = 1
Dim MaxCount
MaxCount = 10
moveTo = CStr(objWshShell.CurrentDirectory) + "\jpg\"
	    WriteFile writePath,"moveTo = " + moveTo
for each file in folder.files
    count = 1
    If (StrComp(fso.GetExtensionName(file.name),"jpg") = 0) Then
	    CheckPath = moveTo + file.name
	    WriteFile writePath,"checkpath = " + checkpath
	    moveBase = CStr(objWshShell.CurrentDirectory) + "\" + Cstr(file.name)
	    
	    flag = IsExistsPath(checkpath)
	    if (not flag) Then	    
	        MoveFileCustom writePath,moveBase,moveTo
	    Else
	        '移動先に同名ファイルがあれば、違う名前にリネームする
	        'リネーム時にファイルフォルダと同じファイル名があれば、違う名前にリネームする
	        ret = FileRenameForMove(MaxCount,moveBase,moveTo)
	    End If
	End If
next 

' サブフォルダ一覧
dim subfolder
'for each subfolder in folder.subfolders
    'msgbox subfolder.name
    'list = list + vbCr  + fso.GetFileName(subfolder)
'next

msgbox "success"
'=======================================================
'移動先に同名ファイルがあれば、違う名前にリネームする
'リネーム時にファイルフォルダと同じファイル名があれば、違う名前にリネームする
Function FileRenameForMove(MaxCount,path,movePath)
	dim ret 'flag
	ret = False
	dim fso
	Set fso = WScript.CreateObject("Scripting.FileSystemObject")
	dim count
	count = 1
	
	'path exists
	If Not fso.FileExists(path) Then 
	    FileRenameForMove = False
	    Exit Function
	End If
	
	dim objfile
	Set objfile = fso.GetFile(path)
	Dim BaseFileName
	BaseFileName = fso.GetBaseName(objFile)

	dim nextname
	Dim CheckPathMoveFolder
    Dim CheckPathNowFolder
    
    CheckPathMoveFolder = movePath & "\" + objFile.Name
	If fso.FileExists(CheckPathMoveFolder) Then
	    '移動先に存在する
	    Dim IsContinue
	    IsContinue = True
	    Do While (IsContinue)
	        '次の名前を取得
	        NextName  = GetNextFileName(fso,objfile,BaseFileName,Count)
	        'msgbox nextname
	        
	        CheckPathMoveFolder = movePath & "\" + NextName
	        CheckPathNowFolder = objFile.ParentFolder & "\" + NextName
	        '移動先に存在せず、今のフォルダにも存在しない、条件両方に合うものリネーム
	        If (Not fso.FileExists(CheckPathMoveFolder)) And (Not fso.FileExists(CheckPathNowFolder)) Then
	            'リネームする
	            'リネーム時にファイルフォルダと同じファイル名があれば、違う名前にリネームする
	            objFile.Name = NextName
	            ret = True
	            Exit Do
	        Else
	            'どちらかに存在する
	            IsContinue = True
	        End If
	        Count = Count + 1
	        If Count > MaxCount Then Exit Do
	    Loop
	Else
	    '移動先に存在しない
	    ret = true '移動する
	    'msgbox ret
	End If

	If Err.Number = 0 Then
	    'WScript.Echo "Success!"
	    ret = True
	Else
	    WScript.Echo "FileRenameForMove Failed : [" & CStr(Err.Number) & "] " & Err.Description
	    ret = False
	End If
	'戻り値
	FileRenameForMove = ret
End Function

'============================================================
Function GetNextFileName(fso,objFile,BaseFileName,ByRef count)
    GetNextFileName = BaseFileName + " (" + CStr(count) + ")." +  fso.GetExtensionName(objFile.Name)
End Function
'===================
Sub WriteFile(path,value)
'msgbox value
dim fso
dim f
If Not (TypeName(path)="String") Then
    msgbox "TypeName(path) not string"
End If
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(path, 8, true)

	If Err.Number > 0 Then
		dim errstr
		errstr = CStr(err.Number) + " : Open Error : " + err.Description
		msgbox errstr
	    WScript.Echo errstr
	Else
	    f.writeline value
	End If

	f.Close
	Set f = Nothing
	Set fso = Nothing
End Sub
'===================
Sub MoveFileCustom(logpath,filepath,movefolder)

dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
fso.MoveFile filepath,movefolder

If Err.Number = 0 Then
    'WScript.Echo "現在のフォルダは " & objWshShell.CurrentDirectory & " です。"
    'WScript.Echo "Success!"
    WriteFile logpath,"move success : " & filepath + " -> " & movefolder
Else
    'WScript.Echo "エラー: " & Err.Description
    WriteFile logpath,"move failed : " & Cstr(Err.Number) & " : " & Err.Description
    WriteFile logpath,"path :" + filepath + " -> " + movefolder
End If
End Sub
'===================
function IsExistsPath(path)
dim ret
dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")
    If objFso.FileExists(checkpath) Then
        'WScript.Echo "ファイルが存在します"
        ret = true
    Else
        ret = false
    End If
Set objFso = Nothing
IsExistsPath = ret
End Function

set file = nothing
set subfolder = nothing
Set objWshShell = Nothing