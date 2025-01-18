'Option Explicit
'On Error Resume Next

Dim  Flag , Tstr , Z
Dim FileName, FullPath


DirPath = GetCurrentDirectory()
FileName = "FunctionList.vbs"
FullPath = DirPath & "\" & FileName '���̓t�@�C����
FileName = GetCurrentFolder & "\" & "MakeFunctionList.txt"   '�o�̓t�@�C����
Tstr = GetFunctionList(FullPath) '�ǂݍ���
Flag = WriteFile(FileName,Tstr)

'FullPath = "D:\zzz\HowTo\Software\Test.vbs"	'���̓t�@�C����
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

	' �J�����g�f�B���N�g���擾.
	Set objShell = CreateObject( "WScript.Shell" )
	curDir = objShell.CurrentDirectory

	Set objShell = Nothing
	GetCurrentDirectory = curDir
End Function
'/////////////////////////////////////////////////////////////////////////
'����������
Function FormInput(Title)
FormInput = InputBox(Title)
'MsgBox (Input & "����͂��܂����B")
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C�����s
Function RunFile (FullPath)
Dim objWshell
Set objWshell = WScript.CreateObject("WScript.Shell")
'�t�@�C�����݃`�F�b�N
objWshell.Run  FullPath
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�H���_���݃`�F�b�N
Function ExistsFolder(Path)
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFso.FolderExists(Path) = True Then
        '���݂��Ă���B
		ExistsFolder = True
    Else
        '���݂��Ă��Ȃ�
		ExistsFolder = False
    End If
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�H���_�R�s�[
Function CopyFolder(CopyPath,PastePath)
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFolder(CopyPath) = True Then
	If ExistsFolder(PastePath) <> True Then
		' �R�s�[��t�H���_�����݂��Ȃ��Ƃ��͍쐬����
		objFSO.CreateFolder(PastePath)
	Else
		CopyFolder = False
	End IF
	WScript.echo Copypath & vbnewline & PastePath
	' �t�H���_�R�s�[
	    objFso.CopyFolder CopyPath, PastePath
		CopyFolder = True
Else
	CopyFolder = False
End If
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C�����݃`�F�b�N
Function ExistsFile (FullPath)
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(FullPath) = True Then
        '���݂��Ă���B
		ExistsFile = True
    Else
        '���݂��Ă��Ȃ�
		ExistsFile = False
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C���R�s�[
Function CopyFile (CopyFullPath , PasteFullPath)
Dim objFso
'Wscript.echo copyfullpath
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFile(CopyFullPath) Then
	 ' �t�@�C���R�s�[
	objFso.CopyFile CopyFullPath, PasteFullPath, True
	CopyFile = True
	Else
		'���݂��Ȃ�
	CopyFile = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�J�����g�t�H���_�擾
Function GetCurrentFolder()
Dim objWshell
Set objWshell = WScript.CreateObject("WScript.Shell")
	GetCurrentFolder = objWshell.CurrentDirectory 
Set objWshell = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�V���[�g�J�b�g���쐬
Function MakeShortCut(BasePath , FileName , MakePath , ShortCutTitle)
Dim objWshell , objShortcut ,BaseFullPath
Set objWshell = WScript.CreateObject("WScript.Shell")
BaseFullPath = BasePath & "\" & FileName
'BasePath �t�H���_���݃`�F�b�N
'BaseFullPath �t�@�C�����݃`�F�b�N
'MakePath �t�H���_���݃`�F�b�N
If Err.Number = 0 Then
    'strDesktopPath = objWshell.SpecialFolders("Desktop")  '�f�X�N�g�b�v�Ɂ@�V���[�g�J�b�g��
    'strWindowsPath = objWshell.ExpandEnvironmentStrings("%WINDIR%") '�������̏ꏊ �V���[�g�J�b�g���̏ꏊ
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
'�t�@�C���ǂݍ���
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
'�t�@�C����������
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
'�������A��

'/////////////////////////////////////////////////////////////////////////
'�e�X�g�t�@�C������C���N���[�h�p�֐����X�g�쐬
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
'�֐����X�g�쐬
Function GetFunctionList(FullPath)
Dim Title , FileName ,FileFullPath , ReadStr ,Aftstr
Title = "�ǂݍ��݃t�@�C��������́i�t���p�X�Łj"
'FullPath = "D:\zzz\Software\Test.vbs"
'FileFullPath = FormInput(Title)
ReadStr = ReadFile(FullPath)
'MsgBox Instr(1,ReadStr,vbNewLine , vbBinaryCompare )
Aftstr = GetFunctionList_MainProcess(ReadStr)
'MsgBox AftStr
GetFunctionList = AftStr
End Function
'/////////////////////////////////////////////////////////////////////////
'�֐����X�g�쐬���C��
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
		'�������烁����
		Redim Preserve Memo(Cnt)
		Memo(Cnt) = n
		Cnt = Cnt + 1
		'���s�{Function ���������A���s�݂̂ƈ�v�����烁������
		Cnt2 = 0
		nbfo1 = InStrRev( Rstr , Kaigyo & "'", n , vbBinaryCompare)
		nbfo2 = InStrRev( Rstr , Kaigyo , n , vbBinaryCompare)
		Do While (nbfo1 = nbfo2) And (nbfo1 > 0) And (nbfo2 > 0) '�������炻�̑O�������i�t�����j�Ɍ���
			'�������������烁����
			tn = nbfo1
			Redim Preserve Memo2(Cnt2)
			Memo2(Cnt2) = nbfo1
			Cnt2 = Cnt2 + 1
			'Memo(Cnt) = nbfo1
			nbfo1 = InStrRev( Rstr , Kaigyo & "'", tn , vbBinaryCompare)
			nbfo2 = InStrRev( Rstr , Kaigyo , tn , vbBinaryCompare)
		Loop
		If Cnt2 > 0 Then '���Ԃ��������Ȃ�̂ňꎞ�ۑ����ĒǋL
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
		'�����������̂����Ƃɂ��̍s�𕶎����
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
	Else 'Function ��������݂��Ȃ�
		GetFunctionList_MainProcess = Tstr
	End If
End If
End Function