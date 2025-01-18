Attribute VB_Name = "Module4"

Type actStruct
    SheetName As String
    cate As String      '���ڃ��[�h
    find As String      '�������[�h
    rcnt As String      'rows�J�E���^
    Ccnt As String      'Cate�J�E���^
End Type



'find_main
'actFind

Type SellRect
    stRow As Integer
    stCol As Integer
    EndRow As Integer
    EndCol As Integer
    tmpRow As Integer
    tmpCol As Integer
    cName As String
    Fflag As Integer
    Num As Integer
    sName As String
End Type
'//////////////////////////////////////////////////////////////////////////////////////////////
Public Function find_main()
    Dim rowcnt As Integer   '�s�J�E���g
    Dim rect2 As SellRect   '�����p
    Dim tmpRange As Range   '�����p
    Dim tmpObj As Object    '�ꎞ
    
    '���ڃ��[�h���ݒ�
    With Worksheets("�ݒ�")
        rowcnt = .Cells(1, 4).Value
        rect2.EndRow = .Cells(65535, 1).End(xlUp).Row
        rect2.EndCol = .Cells(65535, 1).End(xlUp).Column
        rect2.stRow = 1
        rect2.stCol = 1
        Set tmpObj = .Range(.Cells(rect2.stRow, rect2.stCol), .Cells(rect2.EndRow, rect2.EndCol))
        cateCnt = tmpObj.Rows.Count
    End With
    
    '���ڃ��[�h���@�m��
    Dim cateWord() As String
    ReDim cateWord(cateCnt + 2)
    '���ڃ��[�h�ǂݍ��݁@�z���
    cnt = 1
    For Each ForRange In tmpObj
        cateWord(cnt) = ForRange.Value
        cnt = cnt + 1
    Next ForRange


    '�������[�h�錾
    Dim findWord() As String
    ReDim findWord(cateCnt + 1)
    '�������[�h��z���
    cnt = 1
    Set tmpRange = Range(Cells(5, 1), Cells(5, 1 + cateCnt))
    For Each ForRange In tmpRange
        findWord(cnt) = ForRange.Value
        cnt = cnt + 1
    Next ForRange
    

    '�������[�h�\��t���@�������@�f�o�b�N�p
    cnt = 1
    Set tmpRange = Range(Cells(7, 1), Cells(7, 1 + cateCnt))
    For Each ForRange In tmpRange
        ForRange.Value = findWord(cnt)
        cnt = cnt + 1
    Next ForRange
    
    '�������ʂ���������N���A
        '�s�̏I����I��
        '���ׂĂ̏I���s�𒲍�����@��ԑ傫�����̂̍s���ŏI
        rect2.EndRow = 9
        cnt = cateCnt - 2
        For Num = 1 To cnt
            If Cells(65535, Num).End(xlUp).Row > rect2.EndRow Then
                rect2.EndRow = Cells(65535, Num).End(xlUp).Row
            End If
        Next
        Range(Cells(rowcnt - 2, 1), Cells(rect2.EndRow, cateCnt + 1)).Delete

    '�\���̐錾
    Dim actData() As actStruct
    ReDim actData(cateCnt + 1)
    '�����A���ڃ��[�h���󂯓n���p�\���̂�
    For i = 1 To cateCnt
        actData(i).cate = cateWord(i)
        actData(i).find = findWord(i)
    Next i
    actData(0).Ccnt = cateCnt
    '�V�[�g���O�p�z����m�ہ@�ݒ��5��ڂ��V�[�g��
    sCnt = Worksheets("�ݒ�").Cells(65535, 5).End(xlUp).Row
    If sCnt > cateCnt Then
        ReDim Preserve actData(sCnt)
    End If
    '�V�[�g�l�[���ǂݍ��݁@�z���
    With Worksheets("�ݒ�")
        Set tmpRange = .Range(.Cells(1, 5), .Cells(.Cells(65535, 5).End(xlUp).Row, 5))
        cnt = 1
        For Each ForRange In tmpRange
            actData(cnt).SheetName = ForRange.Value
            cnt = cnt + 1
        Next ForRange
    End With

    '�����@�@////////////////////////////////////////////////////////////////////
    cnt = 1
    For cnt = 1 To sCnt
        actData(0).rcnt = rowcnt
        actData(0).SheetName = actData(cnt).SheetName
        rowcnt = actFind(actData())
    Next cnt

    
    
    Worksheets("����").Activate
    Cells(1, 1).Select
    
End Function

'actFind
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
Function actFind(actData() As actStruct) As Integer
    Dim Num As Integer
    Dim rowcnt As Integer
    Dim fRect As SellRect
    Dim tmpRange As Range
    Dim Flag As Integer
    Dim obj As Object
    Dim sCnt As Integer
    '�s�����p��
    rowcnt = actData(0).rcnt
    With Worksheets(actData(0).SheetName)
        .Activate
        '�J�e�S���s�̍ŏ�����Ō�܂őI���@�@///////////////����������W���L��
        fRect.cName = actData(4).cate  ' 1 -> No
        fRect = find_set(fRect)

        '�����͈͂�I��
        Set obj = .Range(.Cells(fRect.stRow, 1), .Cells(fRect.EndRow, 1))
        '///////////////////////////////�@����������W���L���@END
        '///////////////////////////////�@�������C��
        For Each ForRange In obj
            Flag = 0
            For Num = 0 To actData(0).Ccnt
                '�������[�h���@��Ȃ�Ζ���
                If Not (actData(Num + 1).find = "") Then
                    '��v�����̂�����΂P�@��v���Ȃ���΂O�@�啶���E����������ʂ��Ȃ�
                    If InStr(1, ForRange.Offset(0, Num).Value, actData(Num + 1).find, vbTextCompare) > 0 Then
                        Flag = 1
                    Else
                        Flag = 0
                        Exit For
                    End If
                End If
            Next Num
            'flag = 1 �Ȃ�y�[�X�g
            If Flag = 1 Then
                ForRange.EntireRow.Copy
                With Worksheets("����")
                    .Rows(rowcnt).PasteSpecial
                    rowcnt = rowcnt + 1
                End With
                
            End If
        Next ForRange
    End With
    actFind = rowcnt
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
