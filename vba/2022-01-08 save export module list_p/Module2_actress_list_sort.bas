Attribute VB_Name = "Module2"
'act_arr()
'act_setting()
'title_arr()
'num_arr()
'genre_arr()
'set_sell()


'/////////////////////////////////////////////////////////////////////////////////////////////////���D�ʕ��ёւ�
Public Function act_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    rect.stRow = rect.stCol = rect.Num = 0
    '�o�����D�@����ԍ���
    rect.cName = "�o�����D"
    rect.Num = 1
    rect = LineCap(rect)
    '�^�C�g���@�����̉E��
    rect.cName = "�^�C�g��"
    rect.Num = 2
    rect = LineCap(rect)
    '���͒l���ׂđI��
    rect.cName = "�o�����D"
    rect = allSelect(rect)
    '�n�_ + 1 ����I�_��I��
    If ActiveSheet.Name = "����" Then
        rect.stRow = Worksheets("�ݒ�").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '���ёւ�
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), key2:=Columns(rect.stCol + 1)
    
    '���ɖ߂�
    '�o�����D�@��4�Ԗڂ�
    rect.cName = "�^�C�g��"
    rect.Num = 6
    rect = LineCap(rect)
    '�^�C�g���@�����̉E��
    rect.cName = "�o�����D"
    rect.Num = 6
    rect = LineCap(rect)
    
    Cells(1, 1).Select
    '�{�^���̃T�C�Y�ƈʒu���Z�b�g
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////������ёւ�
Public Function act_setting()
    Dim rect2 As SellRect
    Dim cateCnt As Integer
    Dim cnt As Integer
    Dim tmpRange As Range
    Dim sName As String
    Dim setcnt As Integer
    Dim tmpObj As Object
    sName = ActiveSheet.Name
    
    '���ڃ��[�h���ݒ�
    With Worksheets("�ݒ�")
        rect2.EndRow = .Cells(65535, 1).End(xlUp).Row
        rect2.EndCol = .Cells(65535, 1).End(xlUp).Column
        rect2.stRow = 1
        rect2.stCol = 1
        Set tmpObj = Range(.Cells(rect2.stRow, rect2.stCol), .Cells(rect2.EndRow, rect2.EndCol))
        cateCnt = tmpObj.Rows.Count
    End With
    
    Dim cateWord() As String
    ReDim cateWord(cateCnt + 2)
    '���ڃ��[�h�ǂݍ���
    With Worksheets("�ݒ�")
        cnt = 1
        'Set tmpRange = Selection
        For Each ForRange In tmpObj
            cateWord(cnt) = ForRange.Value
            cnt = cnt + 1
        Next ForRange
    End With
    '���ёւ�
    cnt = 1
    Worksheets(sName).Activate
    For cnt = 1 To cateCnt
        rect2.cName = cateWord(cnt)
        rect2.Num = cnt
        rect2 = LineCap(rect2)
    Next cnt
    
    Cells(1, 1).Select
    '�{�^���̃T�C�Y�ƈʒu���Z�b�g
    Dim j As Integer
    Dim obj As OLEObject
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////�^�C�g���ʕ��ёւ�
Public Function title_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    '�^�C�g���@�����̉E��
    rect.cName = "�^�C�g��"
    rect.Num = 1
    rect = LineCap(rect)
    '���͒l���ׂđI��
    rect.cName = "�^�C�g��"
    rect = allSelect(rect)
    '�n�_ + 1 ����I�_��I��
    If ActiveSheet.Name = "����" Then
        rect.stRow = Worksheets("�ݒ�").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '���ёւ��@key2�����D��
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), key2:=Columns(rect.stCol + 4)
    
    '���ɖ߂�
    rect.cName = "�^�C�g��"
    rect.Num = 5
    rect = LineCap(rect)
    Cells(Worksheets("�ݒ�").Cells(1, 4).Value, 1).Select
    Cells(1, 1).Select
'�{�^���̃T�C�Y�ƈʒu���Z�b�g
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////�ԍ������ёւ�
Public Function num_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    '�^�C�g���@�����̉E��
    rect.cName = "No"
    rect.Num = 1
    rect = LineCap(rect)
    '���͒l���ׂđI��
    rect.cName = "No"
    rect = allSelect(rect)
    '�n�_ + 1 ����I�_��I��
    If ActiveSheet.Name = "����" Then
        rect.stRow = Worksheets("�ݒ�").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '���ёւ��@key2�����D��
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), DataOption1:=xlSortTextAsNumbers
    
    '���ɖ߂�
    rect.cName = "No"
    rect.Num = 1
    rect = LineCap(rect)
    Cells(Worksheets("�ݒ�").Cells(1, 4).Value, 1).Select
    Cells(1, 1).Select
'�{�^���̃T�C�Y�ƈʒu���Z�b�g
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
End Function


'/////////////////////////////////////////////////////////////////////////////////////////////////�W�����������ёւ�
Public Function genre_arr()
    Dim obj As Object
    Dim rect As SellRect
    Dim tmpV As Variant
    Dim Num As Integer
    
    rect.stRow = rect.stCol = rect.Num = 0
    '�o�����D�@����ԍ���
    rect.cName = "�ެ��"
    rect.Num = 1
    rect = LineCap(rect)
'    '�^�C�g���@�����̉E��
'    rect.cName = "�^�C�g��"
'    rect.num = 2
'    rect = LineCap(rect)
    '���͒l���ׂđI��
    rect.cName = "�ެ��"
    rect = allSelect(rect)
    '�n�_ + 1 ����I�_��I��
    If ActiveSheet.Name = "����" Then
        rect.stRow = Worksheets("�ݒ�").Cells(1, 4).Value
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    Else
        Set obj = Range(Cells(rect.stRow + 1, rect.stCol), Cells(rect.EndRow, rect.EndCol))
    End If
    '���ёւ�
    obj.Sort _
    Key1:=Cells(rect.stRow, rect.stCol), key2:=Columns(rect.stCol + 1)
    
    '���ɖ߂�
    rect.cName = "�ެ��"
    rect.Num = 8
    rect = LineCap(rect)
    Cells(Worksheets("�ݒ�").Cells(1, 4).Value, 1).Select
    Cells(1, 1).Select
    
    '�{�^���̃T�C�Y�ƈʒu���Z�b�g
    For Each obj In ActiveSheet.OLEObjects
        If obj.Name Like "CommandButton*" Then
            obj.Top = 0
            obj.Left = 300 + j * 70
            obj.Height = 23
            obj.Width = 50
            j = j + 1
        End If
    Next
    
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////���ݒ� set_sell
Public Function set_sell()
    Dim rect2 As SellRect
    Dim cateCnt As Integer
    Dim cnt As Integer
    Dim tmpRange As Range
    Dim sName As String
    Dim setcnt As Integer
    Dim tmpObj As Object

    '���݂̃V�[�g�����L��
    sName = ActiveSheet.Name
    
    '���ڐ�
    cateCnt = Worksheets("�ݒ�").Cells(65535, 1).End(xlUp).Row
    '�����p���ږ��@cateword�@�錾�E�m��
    Dim cateWord() As String
    ReDim cateWord(cateCnt + 2)
    '���ڃ��X�g��z���
    cateWord = word_set()
    
    '���ёւ�
    cnt = 1
    Worksheets(sName).Activate
    For cnt = 1 To cateCnt
        rect2.cName = cateWord(cnt)
        rect2.Num = cnt
        rect2 = LineCap(rect2)
    Next cnt

    'rect2.cName = cateWord(2)
    'rect2 = allSelect2(rect2)
    
    '////////////////////////////////////////////////////////////////////
    '���ݒ�    //////////////////////////////////////////////////////////
    '�ǂݍ���
    setcnt = Worksheets("�ݒ�").Cells(65535, 9).End(xlUp).Row  '�ݒ��9��ڂ̍Ō�̍s�ԍ�
    Dim setVal() As String
    ReDim setVal(setcnt + 1)
    '���ڃ��[�h�ǂݍ���
    With Worksheets("�ݒ�")
        cnt = 1
        '�I��
        Set tmpRange = .Range(.Cells(1, 9), .Cells(setcnt, 9))
        For Each ForRange In tmpRange
            setVal(cnt) = ForRange.Value
            cnt = cnt + 1
        Next ForRange
    End With
    
    'rows�c���ݒ�
    rect2.cName = cateWord(1) ' �^�C�g��
    rect2 = find_set(rect2)
    rect2.EndRow = Cells(65535, rect2.stCol).End(xlUp).Row
    Set tmpObj = Range(Cells(rect2.stRow, rect2.stCol), Cells(rect2.EndRow, rect2.stCol))
    For Each ForRange In tmpObj
'        If forrange.RowHeight > 27 Or forrange.RowHeight < 10 Then
'            forrange.Rows.AutoFit
'        End If
        If ForRange.RowHeight <> 12 Then
            ForRange.RowHeight = 12
        End If
    Next ForRange
    'column���ݒ�
    Cells(1, 1).Select
    For cnt = 1 To setcnt
        Cells(1, cnt).ColumnWidth = setVal(cnt)
    Next cnt
'    '�܂�Ԃ��ĕ\�����Ȃ�
'    For num = 6 To cateCnt + 1
'        rect2.cName = cateWord(num)
'        rect2 = find_set(rect2)
'        Set tmpObj = Cells(1, num).EntireColumn
'        If tmpObj.WrapText = True Then
'            tmpObj.WrapText = False
'        End If
'    Next num
'�{�^���̃T�C�Y�ƈʒu���Z�b�g
    Dim j As Integer
    Dim obj As OLEObject
For Each obj In ActiveSheet.OLEObjects
    If obj.Name Like "CommandButton*" Then
        obj.Top = 0
        obj.Left = 300 + j * 70
        obj.Height = 23
        obj.Width = 50
        j = j + 1
    End If
Next

End Function
