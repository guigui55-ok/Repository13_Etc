Attribute VB_Name = "Ad_Mdl"


'########################################################################################
'Globals
'########################################################################################
Dim gConst As ConstFindSystem

'�O���[�o���ϐ�
'�e�[�u���̃L���v�V�����̃A�h���X���X�g
Public gCaptionAddressList() As String
'�V�[�g�����X�g
Public gSheetList() As String
'�t�B�[���h�����X�g
Public gFielsNameList() As String
'�������ʃ��X�g
Public gResultData() As String
'�f�o�b�O���[�h
'ON�̎��́A���O�Ȃǂ��o�͂���
Dim gDebugMode As Integer
'���O�o�͗p�V�[�g���̐ݒ�
Dim gLogoutSheetName As Integer
'���O�o�͗p�Z���A�h���X�̐ݒ�
Dim gLogoutBeginAddress As String
'���O�o�͗p�A�h���X
'���ݏo�͂��Ă���ꏊ
Dim gLogoutAddress As String
'���O�o�͗p�t�@�C���p�X
Dim gLogoutPath As String
'���O�o�͗p�C���f���g
Dim gLogIndent As Integer
'gLogIndent = 0

Sub CountUpLogIndent(Optional n As Integer = 0)
    gLogIndent = gLogIndent + n
End Sub
Sub CountDownLogIndent(Optional n As Integer = 0)
    gLogIndent = gLogIndent - n
    If gLogIndent < 0 Then
        gLogIndent = 0
    End If
End Sub

'########################################################################################
'Constants
'########################################################################################
Const DEBUG_ON As Integer = 1
Const LOG_TO_IMMIDIATE As Integer = 2
Const LOG_TO_CELL As Integer = 4
Const LOG_TO_FILE As Integer = 8
Const SHOW_ERROR_MSG_BOX As Integer = 1
Const SHOW_ERROR_DEBUG_PRINT As Integer = 2

'########################################################################################
'CommonFunction Module
'########################################################################################
'�f�o�b�O���[�hON�̂Ƃ��Z���֏o�͂���
Sub Logout(Value As Variant)
On Error GoTo ErrRtn
    'OFF�Ȃ�I������
    'DEBUG_MODE_ON=1
    If gDebugMode < 1 Then
        Exit Sub
    End If
    Dim buf As String
    buf = CnvVarToStr(Value)
    buf = Str(Now()) + " " + buf
    '�C�~�f�B�G�C�g�֏o�͂���
    'LOG_TO_IMMIDIATE=2
    If gDebugMode And 2 Then
        ShowDebugPrint buf
    End If
    '�Z���ɏo�͂���
    'LOG_TO_CELL=4
    If gDebugMode And 4 Then
        'gLogoutPath
        Sheets(gLogoutSheetName).Range(gLogoutAddress).Value = buf
    End If
    '�t�@�C���ɏo�͂���
    'LOG_TO_FILE=8
    If gDebugMode And 8 Then
        'gLogoutPath
        Flag = WriteFile(gLogoutPath, buf)
    End If
Exit Sub
ErrRtn:
    Debug.Print ("Module:Ad_Mdl,Function:Logout,Err=" + Err.Number + ":" + Err.Description)
End Sub


'�f�o�b�O���[�hON�̂Ƃ��C�~�f�B�G�C�g�֏o�͂���
Sub ShowDebugPrint(Value As String)
    If gDebugMode >= DEBUG_ON Then
        Debug.Print (Value)
    End If
End Sub

'���b�Z�[�W�{�b�N�X����\������A�f�o�b�O���[�hON�̂Ƃ��C�~�f�B�G�C�g�ւ��o�͂���
Sub ShowDebugMsgBox(Value As String)
    MsgBox Value
    If gDebugMode >= DEBUG_ON Then
        Debug.Print (Value)
    End If
End Sub


'########################################################################################
'�e�[�u����̃f�[�^����������V�X�e�� ���C�����s�֐�
'########################################################################################
'########################################################################################
'�����̌����e�[�u�����當�������������
'���s���\�b�h
Sub FindStringOfMultiTable(Optional debugMode As Integer = 0)
    Dim cFindString As AdCl_FindStringMain
    Set cFindString = New AdCl_FindStringMain
    
    '���s�J�n���̏������O�֏o�͂���
    Debug.Print ("Module:Ad_Mdl , Function:FindStringOfMultiTable , DebugMode:" + Str(debugMode))
    '�f�o�b�O���[�h�ϐ����O���[�o���ƃN���X�����o�֊i�[����
    gDebugMode = debugMode
    cFindString.debugMode = debugMode
    '�@�\�����s����
    Call cFindString.Main
    
    Set cFindString = Nothing
End Sub

