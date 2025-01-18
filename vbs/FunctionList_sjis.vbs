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
DPErr("GetListFile")
End Function
'/////////////////////////////////////////////////////////////////////////
'�z��𕶎���� VbNewLine�ŋ�؂�
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
'�z��𕶎���� VbNewLine�ŋ�؂�
Function AryJoin(ArrayString,delimita)		'As String
	Dim Tstr
	If AryCheckZero(ArrayString) Then
		Tstr = ""
		For i = 0 to Ubound(ArrayString)
			Tstr = Tstr & ArrayString(i) & delimita
		Next
		AryJoin =  Tstr
	Else
		AryJoin = ""
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
End Function
'/////////////////////////////////////////////////////////////////////////
'�t�H���_�R�s�[
Function CopyFolder(CopyPath,PastePath)	'As Boolean
Dim objFso
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
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
		'MsgBox FullPath  & " �����݂��Ă��܂���B" ''�I�����Ă��Ȃ�������^�̒萔�ł��B
    End If
Set objFso = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////
Function DeleteFile(Path)'As Boolean
	'�t�@�C���V�X�e���I�u�W�F�N�g�쐬
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'�t�@�C�����폜����
	objFso.DeleteFile Path
	DPErr("DeleteFile")
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
	'��3����False�̏ꍇ�́A�R�s�[��ɓ����t�@�C�������݂���Ƃ��G���[�ƂȂ�
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
	dim RetStr
	RetStr = ""
	For Each OSInfo In OSInfoCollection '���[�v�̓R���N�V������1�Ԗڂ��Q�Ƃ��邽�߂݂̖̂���
		RetStr = RetStr + "�I�y���[�e�B���O�V�X�e���F" & OSInfo.Caption + vbNewLine
	'    WScript.Echo "�I�y���[�e�B���O�V�X�e���F" & OSInfo.Caption
		RetStr = RetStr + "�o�[�W�����F" & OSInfo.Version + vbNewLine
	'    WScript.Echo "�o�[�W�����F" & OSInfo.Version
		RetStr = RetStr + "�T�[�r�X�p�b�N�F" & OSInfo.CSDVersion + vbNewLine
	'    WScript.Echo "�T�[�r�X�p�b�N�F" & OSInfo.CSDVersion
		RetStr = RetStr + "�V�X�e���f�B���N�g���F" & OSInfo.SystemDirectory + vbNewLine
	'    WScript.Echo "�V�X�e���f�B���N�g���F" & OSInfo.SystemDirectory
		RetStr = RetStr + "�V�X�e���h���C�u�F" & OSInfo.SystemDrive + vbNewLine
	'    WScript.Echo "�V�X�e���h���C�u�F" & OSInfo.SystemDrive
		RetStr = RetStr + "���z�������e�ʁF" & OSInfo.TotalVirtualMemorySize & "Bytes" + vbNewLine
	'    WScript.Echo "���z�������e�ʁF" & OSInfo.TotalVirtualMemorySize & "Bytes"
		RetStr = RetStr + "�����������e�ʁF" & OSInfo.TotalVisibleMemorySize & "Bytes" + vbNewLine
	'    WScript.Echo "�����������e�ʁF" & OSInfo.TotalVisibleMemorySize & "Bytes"
		GetOSName = OSInfo.Caption
	Next
	GetOSName = RetStr
End Function
'/////////////////////////////////////////////////////////////////////////
'�V���[�g�J�b�g���쐬
Function MakeShortCut(BasePath , FileName , MakePath , ShortCutTitle)	'As Boolean
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
'�t�@�C����������
Function WriteFile (FullPath,Wstr)	'As Boolean
Dim ForReading:ForReading = 1
Dim ForWriting:ForWriting = 2
Dim ForAppending:ForAppending = 8
Dim objFso, objFile
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
'MsgBox Wstr
If ExistsFile(FullPath) Then
    If Err.Number = 0 Then
		Set objFile = objFSO.OpenTextFile(FullPath, ForAppending, TristateFalse)
        objFile.WriteLine(Wstr)
		objFile.Close
		WriteFile = True
	Else
		tmp = DPErrIn( "WriteFile", Err.Number, Err.Description)
		WriteFile = False
	End If
Else
    'MsgBox "Path Not Exists[" + FullPath + "]"	
	Set objFile = objFSO.OpenTextFile(FullPath, 2,True)
	objFile.WriteLine(Wstr)
	objFile.Close
	WriteFile = True
	'WriteFile = False
End If
Set objFile = Nothing
Set objFSO = Nothing
DPErr("WriteFile")
End Function
'/////////////////////////////////////////////////////////////////////////
'�������A��

'/////////////////////////////////////////////////////////////////////////
'�֐����X�g�쐬
Function GetFunctionList(FullPath)'As String
	Dim Title , FileName ,FileFullPath , ReadStr ,Aftstr
	'Title = "�ǂݍ��݃t�@�C��������́i�t���p�X�Łj"
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
'//////////////////////////////////////////////////////////////////////////
Function DPErrIn(FuncName,ErrNumber,ErrDescription)
    Dim msg
    If ErrNumber <> 0 Then
        msg = "Error : " & ErrNumber & " : " & ErrDescription
        msg = msg + " , Function = " & FuncName
        'msg = msg + " , Source = " & Err.Source
        'msg = msg + " , Erl = " & CStr(Erl)
        msgbox msg
    End If
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
Function GetDateStr()
	Dim lngYear
	lngYear = Year(Now)
	Dim lngMonth
	lngMonth = Month(Now)
	Dim lngDay
	lngDay = Day(Now)
	Dim lngHour
	lngHour = Hour(Now)
	Dim lngMinute
	lngMinute = Minute(Now)
	Dim lngSecond
	lngSecond = Second(Now)
	Dim RetStr
	RetStr = ""
	RetStr = Right(CStr(lngYear),2) + CStr(lngMonth) + CStr(lngDay) + "_" + _
		CStr(lngHour) + CStr(lngMinute) + CStr(lngMinute)
	DPErr("GetDateStr")
	GetDateStr = RetStr
End Function