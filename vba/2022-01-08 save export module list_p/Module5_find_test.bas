Attribute VB_Name = "Module5"
'find_set


'//////////////////////////////////////////////////////////////////////////////////////////////////////// find_set

Public Function find_set(ByRef FindRect As SellRect) As SellRect
    Dim findObj As Object
    
    '�l����v����Z����T��
    Set findObj = Cells.find(FindRect.cName, lookat:=xlWhole)
    If findObj Is Nothing Then
        '�Ȃ���΃��b�Z�[�W�{�b�N�X�\���@���@�t���O�[���ɂ���
        MsgBox FindRect.cName & "Not Found"
        FindRect.Fflag = 0
        Exit Function
    Else
        '�t���O���R�Ȃ�΁@�����l�̃Z���̈ʒu���L�����ā@�t���O���P�ɂ���
        If FindRect.Fflag = 3 Then
            FindRect.tmpRow = findObj.Row
            FindRect.tmpCol = findObj.Column
            FindRect.Fflag = 1
            find_set = FindRect
            Exit Function
        End If
        '�t���O���P�Ȃ�΁@�����l�̃Z���̈ʒu�Ƃ��̍ŉ��s�̈ʒu���L������
        FindRect.Fflag = 1
        FindRect.stRow = findObj.Row
        FindRect.stCol = findObj.Column
        FindRect.EndRow = Cells(65535, FindRect.stCol).End(xlUp).Row
        FindRect.EndCol = Cells(65535, FindRect.stCol).End(xlUp).Column
        'Cells(FindRect.stRow, FindRect.stCol).Select
    End If
    'findObj.Select
    find_set = FindRect
End Function
'//////////////////////
