Attribute VB_Name = "Module8"
'Option Compare Database 'error
Option Explicit
  
Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
       (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
  
Private Const WM_VSCROLL    As Long = &H115&  ' �����X�N���[��
Private Const SB_LINEUP     As Long = &H0&    ' �����
Private Const SB_LINEDOWN   As Long = &H1&    ' ������
  
' [ACC2002] �X�N���[������Ɛ擪���R�[�h���\������Ȃ�
' http://support.microsoft.com/kb/418706/ja
' ��L���ۂ̑΍�Ƃ��āA�}�E�X�z�C�[���g�p���C�x���g���g���T���v��
Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
  
    ' ���ɃX�N���[�������ꍇ�͊֌W�Ȃ��̂ŁA�I��
    If Count > 0 Then Exit Sub
  
    Dim MaxVisibleRecordCount   As Integer  ' 1 ��ʂɕ\���\�ȍő僌�R�[�h��
    Dim iDetailsH               As Integer  ' �ڍ׃Z�N�V�����S�̂̍���
    Dim iDetailH                As Integer  ' �ڍ׃Z�N�V��������̍���
    Dim iHeaderH                As Integer  ' �t�H�[���w�b�_�[�̍���
    Dim iFooterH                As Integer  ' �t�H�[���t�b�^�[�̍���
    Dim iCurrAboveRows          As Integer  ' �J�����g�s�̏�ɕ\������Ă���s��
  
On Error Resume Next    ' �t�H�[���w�b�_�[/�t�b�^�[�����݂��Ȃ��ꍇ�̃G���[�𖳎�
    iHeaderH = Me.Section(acHeader).Height
    iFooterH = Me.Section(acFooter).Height
On Error GoTo erh
    iDetailH = Me.Section(acDetail).Height
    iDetailsH = Me.InsideHeight - iHeaderH - iFooterH
    MaxVisibleRecordCount = iDetailsH \ iDetailH
    ' �ǉ���������Ă���ꍇ�́A�V�K���R�[�h 1 �s����␳
    If Me.AllowAdditions Then MaxVisibleRecordCount = MaxVisibleRecordCount - 1
  
    ' �S���R�[�h�� 1 ��ʂɕ\���\�ȏꍇ
    ' -- ���ۂɂ́AMaxVisibleRecordCount = Me.Recordset.RecordCount ��
    '    �ꍇ�A�z�C�[���X�N���[���� 1 �s�ڂ�\���ł��邱�Ƃ�����܂����A
    '    ���E���肪����Ȃ��߁A���l�̏ꍇ���܂߂đ΍􂵂܂��B
    If MaxVisibleRecordCount >= Me.Recordset.RecordCount Then
        iCurrAboveRows = (Me.CurrentSectionTop - iHeaderH) / iDetailH
  
        ' ��ɉB��Ă����\�����R�[�h�� 1 �s�ȉ��̏ꍇ(�o�O�ɊY��)
        If (Me.CurrentRecord - iCurrAboveRows - 1) <= 1 Then
            SendMessage Me.hWnd, WM_VSCROLL, SB_LINEUP, 0&  ' �����X�N���[��!
        End If
    End If
    Exit Sub
  
erh:
    MsgBox Err.Description, vbCritical, "���s���G���[" & Err.Number
End Sub

