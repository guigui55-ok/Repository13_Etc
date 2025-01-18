'Start Function

'/////////////////////////////////////////////////////////////////////////
'�f�X�N�g�b�v�p�X
Function GetDeskTopPath()
Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
GetDeskTopPath = objWShell.SpecialFolders("Desktop")
Set objWshell = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C���폜
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
'�t�H���_�폜
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
'�t�@�C���̈ړ�
Function MoveFileF(BaseFullPath,MoveFullPath) 'As Boolean
Dim objFso , tFlag
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	If ExistsFile(BaseFullPath) Then
		If ExistsFile(MoveFullPath) Then
			'�ړ���ɂ��łɑ��݂���
			objFso.DeleteFile MoveFullPath, True
			If Err.Number = 0 Then
				'�폜���ꂽ
				objFso.CopyFile BaseFullPath, MoveFullPath
				objFso.DeleteFile BaseFullPath, True
				MoveFileF = True
			Else
				'�폜���ꂸ
				MoveFileF = False
			End If
		Else
			objFso.CopyFile BaseFullPath, MoveFullPath 
			objFso.DeleteFile BaseFullPath, True
			MoveFileF = True
		End IF
	Else
		'�t�@�C���A�t�H���_�����݂��Ȃ�
		MoveFileF = False
	End If
Else
	MoveFileF = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�H���_�̈ړ� �i�ړ���ɑ��݂�����A��폜->��փR�s�[->���폜 �A���݂��Ȃ���� ��փR�s�[->���폜�j
Function MoveFolderF(BaseFullPath,MoveFullPath) 'As Boolean
Dim objFso , tFlag , objFso2
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	If ExistsFolder(BaseFullPath) Then
		If ExistsFolder(MoveFullPath) Then
			'�ړ���ɂ��łɑ��݂���
			objFso.DeleteFolder MoveFullPath, True
			If Err.Number = 0 Then
				'�폜���ꂽ
				objFso.CopyFolder BaseFullPath, MoveFullPath
				objFso.DeleteFolder BaseFullPath, True
				MoveFolderF = True
			Else
				'�폜���ꂸ
				MoveFolderF = False
			End If
		Else
			objFso.CopyFolder BaseFullPath, MoveFullPath 
			objFso.DeleteFolder BaseFullPath, True
			MoveFolderF = True
		End IF
	Else
		'�t�@�C���A�t�H���_�����݂��Ȃ�
		MoveFolderF = False
	End If
Else
	MoveFolderF = False
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C�����l�[�� 
Function RenameFile(Path , OldName , NewName) 'As Boolean
Dim objFso , FullPathOld , FullPathNew ,objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	FullPathOld = Path & "\" & OldName
	If ExistsFile(FullPathOld) Then
		FullPathNew = Path & "\" & NewName
		If ExistsFile(FullPathNew) Then
			'�ύX��t�@�C�������ɑ��݂���
			'MsgBox FullPathNew & " = True"
			RenameFile = False
		Else
			'�ύX��t�@�C�������݂��Ȃ� �ꍇ���l�[��
			Set objFile = objFso.GetFile(FullPathOld)
			'MsgBox "objFile.Name = " & objFile.Name
			objFile.Name = NewName
			RenameFile = True
		End If
	Else
		'�ύX���t�@�C�������݂��Ȃ�
		RenameFile = False
	End If
End If
Set objFso = Nothing
End Function
'///////////////////////////////////////////////////////////////////////// 
'�G���[�o��
Function OutError(en , es) 'As 
	Dim FileName , Path , Flag
	Tstr = "Err.Number = " & en & vbNewLine & "Err.Description = " & es
	FileName = "ErrorLog.txt"
	Path = GetCurrentFolder & "\" & Filename
	Flag = WriteFile(Path,Tstr)
End Function
'///////////////////////////////////////////////////////////////////////// 
'�t�H���_���X�g�擾
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
'�t�@�C�����X�g�擾
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
'�z��𕶎���Ɂ@VbNewLine�ŋ�؂�
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
'�z����o��
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
'�z�񂪂��邩�`�F�b�N
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
'����������
Function FormInput(Title)	'As 
FormInput = InputBox(Title)
'MsgBox (Input & "����͂��܂����B")
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C�����s
Function RunFile (FullPath)	'As
'On Error Resume Next
Dim objWshell ,Z
If Err.Number = 0 Then
	If ExistsFile(FullPath) Then
		Set objWshell = WScript.CreateObject("WScript.Shell")
		'�t�@�C�����݃`�F�b�N
		objWshell.Run FullPath, vbNormalFocus,True
	Else
		MsgBox "File Not Foune." & vbnewline & FullPath
	End If
Else
End if
Set objWshell = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C�����s
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
'�t�H���_���݃`�F�b�N
Function ExistsFolder(Path)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFso.FolderExists(Path) = True Then
        '���݂��Ă���B
		ExistsFolder = True
    Else
        '���݂��Ă��Ȃ�
		ExistsFolder = False
		'MsgBox Path  & " �����݂��Ă��܂���B"
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�H���_�R�s�[
Function CopyFolder(CopyPath,PastePath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
	If ExistsFolder(CopyPath) = True Then
		If ExistsFolder(PastePath) <> True Then
			' �R�s�[��t�H���_�����݂��Ȃ��Ƃ��͍쐬����
			objFSO.CreateFolder(PastePath)
		Else
			CopyFolder = False
		End IF
		'WScript.echo Copypath & vbnewline & PastePath
		' �t�H���_�R�s�[
		    objFso.CopyFolder CopyPath, PastePath
			CopyFolder = True
	Else
		CopyFolder = False
			'MsgBox CopyPath  & " �����݂��Ă��܂���B"
	End If
End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C�����݃`�F�b�N
Function ExistsFile (FullPath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(FullPath) = True Then
        '���݂��Ă���B
		ExistsFile = True
    Else
        '���݂��Ă��Ȃ�
		ExistsFile = False
		'MsgBox FullPath  & " �����݂��Ă��܂���B"
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C���R�s�[
Function CopyFile (CopyFullPath , PasteFullPath)	'As Boolean
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
Function GetCurrentFolder()	'As String
Dim objWshell
Set objWshell = WScript.CreateObject("WScript.Shell")
	GetCurrentFolder = objWshell.CurrentDirectory 
Set objWshell = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'OS�̎�ނ̖��O���擾
Function GetOSName()	'As String
Dim OSInfoCollection
Set OSInfoCollection = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
For Each OSInfo In OSInfoCollection '���[�v�̓R���N�V������1�Ԗڂ��Q�Ƃ��邽�߂݂̖̂���
'    WScript.Echo "�I�y���[�e�B���O�V�X�e���F" & OSInfo.Caption
'    WScript.Echo "�o�[�W�����F" & OSInfo.Version
'    WScript.Echo "�T�[�r�X�p�b�N�F" & OSInfo.CSDVersion
'    WScript.Echo "�V�X�e���f�B���N�g���F" & OSInfo.SystemDirectory
'    WScript.Echo "�V�X�e���h���C�u�F" & OSInfo.SystemDrive
'    WScript.Echo "���z�������e�ʁF" & OSInfo.TotalVirtualMemorySize & "Bytes"
'    WScript.Echo "�����������e�ʁF" & OSInfo.TotalVisibleMemorySize & "Bytes"
	GetOSName = OSInfo.Caption
Next
End Function
'/////////////////////////////////////////////////////////////////////////
'GetOSName �̖��O���ȗ��� Select Case�p
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
'�V���[�g�J�b�g���쐬
Function MakeShortCut(BasePath , FileName , MakePath , ShortCutTitle,IconFilePath,IconNum)	'As Boolean
Dim objWshell , objShortcut ,BaseFullPath
Set objWshell = WScript.CreateObject("WScript.Shell")
BaseFullPath = BasePath & "\" & FileName
'BasePath �t�H���_���݃`�F�b�N
'BaseFullPath �t�@�C�����݃`�F�b�N
'MakePath �t�H���_���݃`�F�b�N
If Err.Number = 0 Then
    'strDesktopPath = objWshell.SpecialFolders("Desktop")  '�f�X�N�g�b�v�Ɂ@�V���[�g�J�b�g��
    'strWindowsPath = objWshell.ExpandEnvironmentStrings("%WINDIR%") '�������̏ꏊ �V���[�g�J�b�g���̏ꏊ
    Set objShortcut = objWshell.CreateShortcut(MakePath & "\" & ShortCutTitle & ".lnk") '�V���[�g�J�b�g�쐬��
    objShortcut.Description = ShortCutTitle
'    objShortcut.HotKey = "CTRL+ALT+N"
    objShortcut.IconLocation = IconFilePath & "," & IconNum
    objShortcut.TargetPath = BasePath & "\" & FileName	'�V���[�g�J�b�g�쐬��
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
'�t�@�C���ǂݍ���
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
'�t�@�C����������
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
'�������A��

'/////////////////////////////////////////////////////////////////////////
'�֐����X�g�쐬
Function GetFunctionList(FullPath)	'As String
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
'/////////////////////////////////////////////////////////////////////////
'����e�X�g
Function Testmsg() 'As
	MsgBox "Test"
End Function


'/////////////////////////////////////////////////////////////////////////
'�t�@�C�����݃`�F�b�N
Function ExistsFileL(FullPath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(FullPath) = True Then
        '���݂��Ă���B
		ExistsFileL = True
    Else
        '���݂��Ă��Ȃ�
		ExistsFileL = False
		'MsgBox FullPath  & " �����݂��Ă��܂���B"
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�@�C���������݁@�Ȃ���΍��
Function WriteMakeTextL(FullPath,Wstr) 'As Boolean
Dim objFso , objStm , Rstr , tFlag
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
If ExistsFileL(FullPath) Then
	'���݂���ꍇ�ǂݍ���ŏ�������
	Rstr = ReadFileL(FullPath)
	Rstr = Rstr & Vbnewline & Wstr
	tFlag = WriteFileL(FullPath,Rstr)
Else
	'���݂��Ȃ��ꍇ�͍���ď�������
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
'�t�@�C����������
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
'�t�@�C���ǂݍ���
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
