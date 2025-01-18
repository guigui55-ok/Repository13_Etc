Attribute VB_Name = "CmMd_Base"

'//////////////////////////////////////////////////////////////////
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type


Public Enum ComponentType
    STANDARD_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM = 3
End Enum

'////////////////////////////////////////////////////////////////////////////
Function GetGlobalMemoryStatus(KindString As String) As Long
    '===========================
    Dim MemoryData As MEMORYSTATUS
    Dim ResultLong As Long
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    GlobalMemoryStatus MemoryData
    With MemoryData
        Select Case KindString
            Case "�����������T�C�Y", "dwTotalPhys": ResultLong = .dwTotalPhys
            Case "�g�p�\�ȕ���������", "dwAvailPhys": ResultLong = .dwAvailPhys
            Case "�y�[�W�t�@�C���T�C�Y", "dwTotalPageFile": ResultLong = .dwTotalPageFile
            Case "�g�p�\�ȃy�[�W�t�@�C��", "dwAvailPageFile": ResultLong = .dwAvailPageFile
            Case "���z�������T�C�Y", "dwTotalVirtual": ResultLong = .dwTotalVirtual
            Case "�g�p�\�ȉ��z������", "dwAvailVirtual": ResultLong = .dwAvailVirtual
            Case Else
                GoTo ErrRtn
        End Select
        'Format(.dwAvailVirtual / 1024, "#,##0") & "KB"
    End With
    GetGlobalMemoryStatus = ResultLong
    '==========  End  ==========
Exit Function
ErrRtn:
DPErr
GetGlobalMemoryStatus = -2
End Function
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'�G���[�o��
Public Sub ErrDebugPrintOut()
    Debug.Print "[Err.Num = " & Err.Number & "] Err.Msg = " & Err.Description
End Sub
'////////////////////////////////////////////////////////////////////////////
'���̃N���X�ȂǂɈڂ�ƃV�X�e���G���[�������Ă��܂��̂Ń�������
Sub SaveSystemErrorByModule( _
        SystemErrorDesctiption As String, _
        ByRef ErrorDescriptionArrayString() As String, _
        SystemErrorNumberInteger As Integer, _
        ByRef ErrorNumberString As String, _
        Optional FunctionName As String _
        )
    Dim cTempArray As New Cl_Array
    Dim AddString As String
On Error GoTo ErrRtn
    
    AddString = FunctionName & "[" & SystemErrorNumberInteger & "]" & SystemErrorDesctiption
    ErrorDescriptionArrayString = cTempArray.GetValueDirect _
        .StringAddValueLastElements_RtnString( _
            ErrorDescriptionArrayString, AddString _
    )
'    DP (ErrorDescriptionArrayString): Stop
    ErrorNumberString = ErrorNumberString & "S" & SystemErrorNumberInteger
    
    
Exit Sub
ErrRtn:
    DPErr
End Sub
'////////////////////////////////////////////////////////////////////////////
'���̃N���X�ȂǂɈڂ�ƃV�X�e���G���[�������Ă��܂��̂Ń�������
'���N���X�͕ʊ֐� ������
'Err.Desctiption,Err.Number�A�֐����@�ɕ����Ă���
Sub SaveSystemErrorByModuleForBaseClass( _
        ByRef ErrorDescriptionArrayString() As String, _
        SystemErrorDesctiption As String, _
        ByRef ErrorNumberArrayInteger() As Integer, _
        SystemErrorNumberInteger As Integer, _
        ByRef FunctionNameArrayString() As String, _
        Optional FunctionName As String _
        )
    Dim AddString As String
On Error GoTo ErrRtn
    'Desctiption
    AddString = SystemErrorDesctiption
    If IsArrayExists(ErrorDescriptionArrayString) Then
        'Exists True
        ReDim Preserve ErrorDescriptionArrayString(UBound(ErrorDescriptionArrayString) + 1)
        ErrorDescriptionArrayString(UBound(ErrorDescriptionArrayString)) _
            = AddString
    Else
        'exists False
        ReDim ErrorDescriptionArrayString(0)
        ErrorDescriptionArrayString(0) = AddString
    End If
    
    'Number
    If IsArrayExists(ErrorNumberArrayInteger) Then
        'Exists True
        ReDim Preserve ErrorNumberArrayInteger(UBound(ErrorNumberArrayInteger) + 1)
        ErrorNumberArrayInteger(UBound(ErrorNumberArrayInteger)) _
            = SystemErrorNumberInteger
    Else
        'exists False
        ReDim ErrorDescriptionArrayString(0)
        ErrorDescriptionArrayString(0) = AddString
    End If
    
    'Function Name
    AddString = FunctionName
    If IsArrayExists(FunctionNameArrayString) Then
        'Exists True
        ReDim Preserve FunctionNameArrayString(UBound(FunctionNameArrayString) + 1)
        FunctionNameArrayString(UBound(FunctionNameArrayString)) _
            = AddString
    Else
        'exists False
        ReDim FunctionNameArrayString(0)
        FunctionNameArrayString(0) = AddString
    End If
    
    
Exit Sub
ErrRtn:
    DPErr
End Sub
'////////////////////////////////////////////////////////////////////////////
'�G���[����s���������Ƃ�
Sub DeleteOneLineSystemErrorByModuleForBaseClass( _
        ByRef ErrorDescriptionArrayString() As String, _
        ByRef ErrorNumberArrayInteger() As Integer, _
        ByRef FunctionNameArrayString() As String _
        )
On Error GoTo ErrRtn
    '========== Begin ==========
    If IsArrayExists(ErrorDescriptionArrayString) Then
        Dim CountElement As Long
        CountElement = UBound(ErrorDescriptionArrayString)
        If CountElement = 0 Then
            Erase ErrorDescriptionArrayString
            Erase ErrorNumberArrayInteger
            Erase FunctionNameArrayString
        Else
            ReDim Preserve ErrorDescriptionArrayString(CountElement - 1)
            ReDim Preserve ErrorNumberArrayInteger(CountElement - 1)
            ReDim Preserve FunctionNameArrayString(CountElement - 1)
        End If
    End If
    '==========  End  ==========
Exit Sub
ErrRtn:
    Set cTempArray = Nothing
    DPErr
End Sub
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'@Web
'2015.03.21
'����������������������
Function GetTabDataByUrl_RtnString(UrlString As String) As String
    Dim objIE As Object 'Web�擾�p
    Dim TagData As String
On Error GoTo ErrRtn
'Web�擾
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Navigate URL
    '�y�[�W�̕\��������҂��܂��B
'    While ObjIE.ReadyState <> 4 Or ObjIE.Busy = True '.ReadyState <> 4�̊Ԃ܂��B
'        DoEvents '�d��
'    Wend
    '�_�E�����[�h�҂�
    Do While objIE.Busy
    Loop
    
    TagData = objIE.Document.getElementsByTagName("BODY").Item(0).InnerHTML 'o
    objIE.Quit
'    Debug.Print Right(TagData, 200): Stop
    GetTabDataByUrl_RtnString = TagData
    Set objIE = Nothing
Exit Function
ErrRtn:
'    Call DPErr
    GetTabDataByUrl_RtnString = "__ERROR__"
    Set objIE = Nothing
Exit Function
End Function
'/////////////////////////////////////////////////////////////////////////////
'2015.03.21
Function GetHtml_RtnString(UrlString As String) As String
On Error GoTo ErrRtn
    '===========================
    Dim objITEM As Object 'for each
    Dim objIE As Object
    Dim TagArrayString() As String
    Dim Count As Integer
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
   Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = False
    
    objIE.Navigate UrlString
    Call untilReady(objIE)  '��WAIT
   
    '�Z���N���A
'    Sheets(SheetName).Range("A1").CurrentRegion.Offset(1, 0).ClearContents
    
    Count = 0
    For Each objITEM In objIE.Document.getElementsByTagName("BODY")
        '�������݈ʒu
'        Sheets(SheetName).Cells(j, i) = objITEM.innerText  '�^�O�̓��e�̂�
        ReDim Preserve TagArrayString(Count)
        TagArrayString(Count) = objITEM.InnerHTML
        Count = Count + 1
    Next
    objIE.Quit
    Set objITEM = Nothing
    Set objIE = Nothing

    '==========  End  ==========
GetHtml_RtnString = TagArrayString(0)
Exit Function
ErrRtn:
'    DPErr
    objIE.Quit
    Set objITEM = Nothing
    Set objIE = Nothing
GetHtml_RtnString = TagArrayString(0)
End Function


Sub untilReady(objIE As Object, Optional ByVal WaitTime As Integer = 10)
    Dim starttime As Date
    starttime = Now()
    Do While objIE.Busy = True Or objIE.ReadyState <> READYSTATE_COMPLETE
        DoEvents
        If Now() > DateAdd("S", WaitTime, starttime) Then
            Exit Do
        End If
    Loop
    DoEvents
End Sub
'/////////////////////////////////////////////////////////////////////////////
'Web�擾�A�@�g���Â炢
Function WebGetTest1(SheetName As String, UrlString As String) As Long
Dim NowSheetName As String
On Error GoTo ErrRtn
    '========== Begin ==========
'    UrlString = "http://info.finance.yahoo.co.jp/ranking/?kd=1&tm=d&vl=a&mk=1&p=1"
    NowSheetName = ActiveSheet.Name
    Sheets(SheetName).Activate
    
    '�V�[�g�A�N�e�B�u�ɂ��Ȃ���΂����Ȃ��H
    
    With Sheets(SheetName).QueryTables.Add _
    ( _
        Connection:="URL;" & UrlString _
        , Destination:=Range("$A$1") _
    )
        .Name = "?kd=1&tm=d&vl=a&mk=1&p=1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        
        .WebSelectionType = xlEntirePage
            'xlEntirePage (���ׂ�)
            'xlAllTables (����l / ���ׂẴe�[�u��)
            'xlSpecifiedTables (����̃e�[�u��)
        
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
    Sheets(NowSheetName).Activate
    '==========  End  ==========
WebGetTest1 = 1
Exit Function
ErrRtn:
WebGetTest1 = -1
DPErr
End Function
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'@Cnv
'������z������s�ŋ�؂��ĕ������
Function CnvAryToStr(VarAry As Variant) As String
    '===========================
    Dim RtnStr As String
    Dim i As Integer
On Error GoTo ErrRtn
    '========== Begin ==========
    '�^�`�F�b�N
    If Not (VarType(VarAry) = vbArray + vbString) Then
        If VarType(VarAry) < vbArray Then
            '�z��łȂ���� ������ɂ��Ė߂�
            AryCompPlusLineRtnStr = CStr(VarAry)
            Exit Function
        End If
    End If
    '�z�񑶍݃`�F�b�N
'    If Not IsArrayExists(VarAry) Then GoTo ErrRtn
    If VarType(VarAry) < vbArray Then GoTo ErrRtn
    '����
    RtnStr = ""
    For i = 0 To UBound(VarAry)
        RtnStr = RtnStr & CStr(VarAry(i)) & vbNewLine
    Next i
    '�Ō�͗]�v�Ȃ̂�
    RtnStr = Left(RtnStr, Len(RtnStr) - 1)
    '==========  End  ==========
CnvAryToStr = RtnStr
Exit Function
ErrRtn:
CnvAryToStr = RtnStr
End Function
'////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
Function CnvVarToStr(ArgumentVariant As Variant) As String
    Dim BufferString As String
    Dim BufferVariant As Variant
    Dim BufferArrayString As String
    Dim i As Integer
    '========== Begin ==========
    If VarType(ArgumentVariant) = vbVariant Then Stop
    If VarType(ArgumentVariant) < vbArray Then
        BufferVariant = ArgumentVariant
        Select Case VarType(BufferVariant)
            Case vbBoolean:     BufferString = CStr(BufferVariant)
            Case vbByte:        BufferString = CStr(BufferVariant)
            Case vbCurrency:    BufferString = CStr(BufferVariant)  '�ʉ�
            Case vbDataObject:  BufferString = "Data Is DataObject Type"
            Case vbDecimal:     BufferString = CStr(BufferVariant) '10�i��
            Case vbDouble:      BufferString = CStr(BufferVariant)
            Case vbDate:        BufferString = CStr(BufferVariant)
            Case vbEmpty:       BufferString = "Data Is Empty Type"
            Case vbError:       BufferString = "Data Is Error Type"
            Case vbInteger:     BufferString = CStr(BufferVariant)
            Case vbLong:        BufferString = CStr(BufferVariant)
            Case vbNull:        BufferString = ""
            Case vbObject:      BufferString = "Data Is Object Type"
            Case vbSingle:      BufferString = CStr(BufferVariant)
            Case vbString:      BufferString = CStr(BufferVariant)
            Case vbUserDefinedType:      BufferString = CStr(BufferVariant)
            Case vbVariant:     BufferString = CStr(BufferVariant)
            Case Else: Stop
        End Select
    Else '�z��
        BufferArrayString = ""
        'VarType(Var) >= vbArray
        For i = 0 To UBound(ArgumentVariant)
            BufferVariant = ArgumentVariant(i)
            Select Case VarType(BufferVariant)
                Case vbBoolean:     BufferString = CStr(BufferVariant)
                Case vbByte:        BufferString = CStr(BufferVariant)
                Case vbCurrency:    BufferString = CStr(BufferVariant)  '�ʉ�
                Case vbDataObject:  BufferString = "Data Is DataObject Type"
                Case vbDecimal:     BufferString = CStr(BufferVariant) '10�i��
                Case vbDouble:      BufferString = CStr(BufferVariant)
                Case vbDate:        BufferString = CStr(BufferVariant)
                Case vbEmpty:       BufferString = "Data Is Empty Type"
                Case vbError:       BufferString = "Data Is Error Type"
                Case vbInteger:     BufferString = CStr(BufferVariant)
                Case vbLong:        BufferString = CStr(BufferVariant)
                Case vbNull:        BufferString = ""
                Case vbObject:      BufferString = "Data Is Object Type"
                Case vbSingle:      BufferString = CStr(BufferVariant)
                Case vbString:      BufferString = CStr(BufferVariant)
                Case vbUserDefinedType:      BufferString = CStr(BufferVariant)
                Case vbVariant:     BufferString = CStr(BufferVariant)
                Case Else
                    '�z��̔z��̏ꍇ������B�[�ǂ��͂��Ȃ�
                    If VarType(BufferVariant) > vbArray Then
                        BufferString = "Data is Array of Array"
                    End If
            End Select
            BufferArrayString = BufferArrayString + " , " + BufferString
            If i < UBound(ArgumentVariant) Then
            Else
                '�I�����ɖ߂�l�p�ϐ��ɖ߂�
                BufferString = BufferArrayString
            End If
        Next i
    End If
    CnvVarToStr = BufferString
    '========== End ==========
Exit Function
ErrRtn:
    CnvVarToStr = "System Error : " & Err.Number & _
        " [" & Err.Description & "]"
End Function
'////////////////////////////////////////////////////////////////////////////
'@Cnv
'�t�B�[���h�ǉ��E�ύX�p�@�^�Z�b�g
Function CnvFieldTypeStrToInt(FldType As String) As Integer
    Dim TypeInt As Integer
Select Case FldType
    Case "dbBoolean", "BOOLEAN", "Boolean", "boolean": TypeInt = dbBoolean
    Case "dbByte", "BYTE", "Byte", "byte": TypeInt = dbByte
    Case "dbInteger", "INTEGER", "Integer", "integer": TypeInt = dbInteger
    Case "dbLong", "LONG", "Long", "long": TypeInt = dbLong
    Case "dbCurrency", "CURRENCY", "Currency", "currency": TypeInt = dbCurrency
    Case "dbSingle", "SINGLE", "Single", "single": TypeInt = dbSingle
    Case "dbDouble", "DOUBLE", "Double", "double": TypeInt = dbDouble
    Case "dbDate", "DATE", "Date", "date": TypeInt = dbDate
    Case "dbText", "TEXT", "Text", "text": TypeInt = dbText '10
    Case "dbDate", "DATE", "Date", "date": TypeInt = dbDate
    Case "dbLongBinary", "LONGBINARY", "LongBinary", "longbinary": TypeInt = dbLongBinary
    Case "dbMemo", "MEMO", "Memo", "memo": TypeInt = dbMemo
    Case Else: TypeInt = 10
End Select
CnvFieldTypeStrToInt = TypeInt
End Function
'//////////////////////////////////////////////////////////////////////////
'�����񂪂��ׂĐ����ł���΁A�����^�ɕϊ�
Function CnvIntegerByIsNumeric(ArgumentData As Variant) As Integer
    If IsNumeric(ArgumentData) Then
        CnvIntegerByIsNumeric = CInt(ArgumentData)
    Else
        CnvIntegerByIsNumeric = 0
    End If
End Function
'//////////////////////////////////////////////////////////////////////////
'Boolean�^��CVar CStr �ł�����ł��Ȃ��̂ŕϊ�����
Function CnvBoolToString(FlagBool As Boolean) As String
    If FlagBool Then
        CnvBoolToString = "True"
    Else
        CnvBoolToString = "False"
    End If
End Function
'//////////////////////////////////////////////////////////////////////////
'�����񂪂��ׂĐ����ł���΁A�����^�ɕϊ�
Function CnvStrToInt(data As String) As Integer
    If IsNumeric(data) Then
        CnvStrToInt = CInt(data)
    Else
        CnvStrToInt = 0
    End If
End Function
'//////////////////////////////////////////////////////////////////////////
Function CnvStrToLng(BufStr As String) As Long
    If IsNumeric(BufStr) Then
        CnvStrToLng = CLng(BufStr)
    Else
        CnvStrToLng = 0
    End If
End Function
'//////////////////////////////////////////////////////////////////////////
'Boolean�^��CVar CStr �ł�����ł��Ȃ��̂ŕϊ�����
Function CnvStrToBln(Flag As Boolean) As String
    If Flag Then
        CStrB = "True"
    Else
        CStrB = "False"
    End If
End Function
'//////////////////////////////////////////////////////////////////////////
'�o���A���g�^(�z��łȂ�)�̕�����z����AString(),Integer()�Ȃǂɕϊ�
Function CnvAryVarToAryInt(Ary As Variant) As Integer()
    Dim TIntA() As Integer, i As Integer
On Error GoTo ErrorHandler
    If VarType(Ary) = vbArray + vbInteger Then
        For i = 0 To UBound(Ary)
            ReDim Preserve TIntA(i)
            TIntA(i) = CInt(Ary(i))
        Next i
    Else
        ReDim TIntA(0)
    End If
    CnvAryVarToAryInt = TIntA
Exit Function
ErrorHandler:
'    Call DPErr: Stop
    CnvAryVarToAryInt = TIntA
End Function
'//////////////////////////////////////////////////////////////////////////
'�o���A���g�^(�z��łȂ�)�̕�����z����AString()�ɕϊ�
Function CnvAryVarToAryStr(Ary As Variant) As String()
    '===========================
    Dim FlagLng As Long
    Dim BufAryStr() As String
    Dim BufStr As String
On Error GoTo ErrRtn
    '========== Begin ==========
    If VarType(Ary) > vbArray Then
        'Ary�͔z��
        For i = 0 To UBound(Ary)
        '�z��̔z�񂩂�����Ȃ��̂� 2����
        '���̏ꍇ��0�Ԃ̂� "Ary(0):"�����킦��
            If VarType(Ary(i)) > vbArray Then
                BufStr = CnvVarToStr("Ary(0):" & Ary(i)(0))
                BufAryStr = ArrayStringgRedimAndAppendForString_RtnArrayString( _
                    BufAryStr, BufStr)
            Else
                '�P��
                BufStr = CnvVarToStr(Ary(i))
                BufAryStr = ArrayStringgRedimAndAppendForString_RtnArrayString( _
                    BufAryStr, BufStr)
            End If
        Next i
    Else
        'Ary�͒P��
        BufStr = CnvVarToStr(Ary)
        ReDim BufAryStr(0)
        BufAryStr(0) = BufStr
    End If
    '==========  End  ==========
CnvAryVarToAryStr = BufAryStr
Exit Function
ErrRtn: CnvAryVarToAryStr = BufAryStr
End Function
'//////////////////////////////////////////////////////////////////////////
'�o���A���g�^(�z��łȂ�)�̕�����z����AString()�ɕϊ�
Function CnvArrayVariantToArrayString(ArgumentArrayVariant As Variant) As String()
    Dim BufferStringArray() As String
    Dim i As Integer
On Error GoTo ErrRtn
    '========== Begin ==========
    If VarType(ArgumentArrayVariant) = vbArray + vbString Then
        '�󂯎�����z�񂪁@������z��
        For i = 0 To UBound(ArgumentArrayVariant)
            ReDim Preserve BufferStringArray(i)
            BufferStringArray(i) = CStr(ArgumentArrayVariant(i))
        Next i
    Else
        If VarType(ArgumentArrayVariant) = vbArray + vbVariant Then
        '�󂯎�����z�񂪁@�o���A���g�^�z��
            For i = 0 To UBound(ArgumentArrayVariant)
                ReDim Preserve BufferStringArray(i)
                BufferStringArray(i) = CStr(ArgumentArrayVariant(i))
            Next i
        Else
            ReDim BufferStringArray(0)
        End If
'        Debug.Print VarType(Ary)
    End If
    CnvArrayVariantToArrayString = BufferStringArray
Exit Function
    '========== End ==========
ErrRtn:
    CnvArrayVariantToArrayString = BufferStringArray
End Function
'/////////////////////////////////////////////////////////////////////////////
'������z����@�����^�z���
Function CnvAryStrToAryInt(BeforeArrayString() As String) As Integer()
    Dim AfterArrayInteger() As Integer
    CnvAryStrToAryInt = CnvArrayStringToArrayInteger(BeforeArrayString)
End Function
'/////////////////////////////////////////////////////////////////////////////
'������z����@�����^�z���
Function CnvArrayStringToArrayInteger(BeforeArrayString() As String) As Integer()
    Dim AfterArrayInteger() As Integer
On Error GoTo ErrRtn
    '========== Begin ==========
    ReDim AfterArrayInteger(UBound(BeforeArrayString))
    With Block
     Dim i As Integer
        For i = 0 To UBound(BeforeArrayString) - 1
            'IsNumeric�Ŕ��f����
            If IsNumeric(BeforeArrayString(i)) Then
                AfterArrayInteger(i) = CInt(BeforeArrayString(i))
            Else
                AfterArrayInteger(i) = 0
            End If
        Next i
    End With
    CnvArrayStringToArrayInteger = AfterArrayInteger
Exit Function
    '========== End ==========
ErrRtn:
CnvArrayStringToArrayInteger = AfterArrayInteger
End Function
'//////////////////////////////////////////////////////////////////////////
'�ό��̕ϐ��𕶎���z��ɂ���
Function CnvVarToAryForParam(ParamArray Ary() As Variant) As String()
    Dim i As Integer
    Dim TStrA() As String
    Dim cnt As Long
On Error GoTo ErrorHandler
    If VarType(Ary(0)) < vbArray Then
        cnt = 0
        For i = 0 To UBound(Ary)
            If VarType(Ary(i)) < vbArray Then '�z��̉\������
                ReDim Preserve TStrA(cnt)
                TStrA(cnt) = CStr(Ary(i))
                cnt = cnt + 1
            Else
                '�z��̂Ƃ��͋�ɂ��Ă���
                ReDim Preserve TStrA(cnt)
'                TStrA(Cnt) = ""
                cnt = cnt + 1
            End If
        Next i
    Else 'Ary�z��ł͂Ȃ�
        ReDim TStrA(0)
    End If
    CnvVarToAryForParam = TStrA
Exit Function
ErrorHandler:
'    Call DPErr: Stop
    CnvVarToAryForParam = TStrA
End Function

'//////////////////////////////////////////////////////////////////////////
'MsgBox�֐��̖߂�l���ے�I�Ȃ��̂ł���� True
'���֐����FIsFalseForMsgBox
Function CnvFlagOfIntegerToBool(FlagInt As Integer) As Boolean
Dim Flag As Boolean
    Select Case FlagInt
        Case vbOK: Flag = False 'OK
        Case vbCancel: Flag = True '�L�����Z��
        Case vbAbort: Flag = True '���~
        Case vbRetry: Flag = False '�Ď��s
        Case vbIgnore: Flag = False '����
        Case vbYes: Flag = False '�͂�
        Case vbNo: Flag = True '������
    End Select
    CnvFlagOfIntegerToBool = False
End Function
'/////////////////////////////////////////////////////////////////////////////
'�G���[�ŕϐ�����Ƃ��g�p
'Var as Variant ���z��̂Ƃ� debug.print Var �ŃG���[�ɂȂ�̂�
'0�Ԗڂ������݂邽��
Function CnvArrayVariantOnlyTheFirstToString( _
        BeforeArrayVariant As Variant _
        ) As String
    Dim BufferString As String
On Error GoTo ErrRtn
    '========== Begin ==========
    Select Case VarType(BeforeArrayVariant)
        '�z��ł���
        Case Is < vbArray
            BufferString = CStr(BeforeArrayVariant)
        '�z��łȂ�
        Case Is > vbArray
            If IsArrayExists(BeforeArrayVariant) Then
                BufferString = CStr(BeforeArrayVariant(0)) & " (Array)"
            Else
                BufferString = "(Not Array)"
            End If
        '���̑�
        Case Else
            BufferString = "(Not Array)"
    End Select
    '100�����ȏ�͒������Ȃ̂�
    If Len(BufferString) > 100 Then
        BufferString = Left(BufferString, 100) & "......"
    End If
    CnvArrayVariantOnlyTheFirstToString = BufferString
Exit Function
    '========== End ==========
ErrRtn:
    CnvArrayVariantOnlyTheFirstToString = ""
End Function
'//////////////////////////////////////////////////////////////////////////
'$A$1:$E$5  -> $A$1
Function CnvAddressTableToSingleOfStart_RtnAddress(ADTable As String) As String
    '===========================
    Dim RtnAD As String
    Dim BufAryStr() As String
On Error GoTo ErrRtn
    '========== Begin ==========
    BufAryStr = Split(ADTable, ":")
    RtnAD = BufAryStr(0)
    '==========  End  ==========
CnvAddressTableToSingleOfStart_RtnAddress = RtnAD
Exit Function
ErrRtn: CnvAddressTableToSingleOfStart_RtnAddress = ""
End Function
'//////////////////////////////////////////////////////////////////////////
'$A$1:$E$5  -> $E$5
Function CnvAddressTableToSingleOfEnd_RtnAddress(ADTable As String) As String
    '===========================
    Dim RtnAD As String
    Dim BufAryStr() As String
On Error GoTo ErrRtn
    '========== Begin ==========
    BufAryStr = Split(ADTable, ":")
    RtnAD = BufAryStr(1)
    '==========  End  ==========
CnvAddressTableToSingleOfEnd_RtnAddress = RtnAD
Exit Function
ErrRtn: CnvAddressTableToSingleOfEnd_RtnAddress = ""
End Function
'/////////////////////////////////////////////////////////////////////////////
'�A�h���X���X�g����������̃I�t�Z�b�g�l��Value�ɕϊ�
Function CnvAddressListMoveOffsetToValueList( _
        GetSheetName As String, _
        AddressList() As String, _
        OffsetRow As Long, _
        OffsetCol As Long _
    ) As String()
    '===========================
    Dim ValueList() As String
On Error GoTo ErrRtn
    '========== Begin ==========
    If Not IsArrayExists(AddressList) Then GoTo ErrRtn
    If Not IsSheetExists(GetSheetName) Then GoTo ErrRtn
    ReDim ValueList(UBound(AddressList))
    
    '�܂��A�h���X����������
    For i = 0 To UBound(AddressList)
        '�I�t�Z�b�g��}�C�i�X�ɂȂ�Ȃ���
        If IsAddress(AddressList(i)) Then
            If (Range(AddressList(i)).Row + OffsetRow > 0) And _
                (Range(AddressList(i)).Column + OffsetCol > 0) Then
                ValueList(i) = Sheets(GetSheetName).Range(AddressList(i)) _
                    .Offset(OffsetRow, OffsetCol).Value
            Else
                '�I�t�Z�b�g��̃A�h���X���s��
                ValueList(i) = "ERROR!"
            End If
        Else
            '�A�h���X���s��
            ValueList(i) = "ERROR!"
        End If
    Next i
    CnvAddressListMoveOffsetToValueList = ValueList
    '==========  End  ==========
Exit Function
ErrRtn:
CnvAddressListMoveOffsetToValueList = ValueList
End Function
'/////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
' �󂯎�����ϐ��ɂ��āA�^�t���[�ŃC�~�f�B�G�C�g�֏o�͂���
' @DP
Sub DP100(ArgumentVariant As Variant)
        Dim BufferString As String
        Dim BufferVariant As Variant
        Dim i As Integer
    '========== Begin ==========
On Error GoTo ErrRtn
'    Debug.Print VarType(Var) '*
    '����`�ł�������z��́F8200
    'Array 8192
    'String 8
    If VarType(ArgumentVariant) = vbVariant Then Stop
    Select Case VarType(ArgumentVariant)
        Case vbEmpty: Debug.Print "vbEmpty"
        Case Else
            If VarType(ArgumentVariant) < vbArray Then
                '�z��łȂ�
                BufferString = CnvVarToStr(ArgumentVariant) '�ϊ�
                Debug.Print Left(BufferString, 100)
            Else
                '�z��
                'VarType(Var) >= vbArray
                For i = 0 To UBound(ArgumentVariant)
                    BufferVariant = ArgumentVariant(i)
                    BufferString = CnvVarToStr(BufferVariant) '�ϊ�
                    Debug.Print format(i, "000") & ":" & Left(BufferString, 100)
                Next i
            End If
    End Select
Exit Sub
    '========== End =========
ErrRtn:
    Select Case Err.Number
        Case 9
            Debug.Print "[Ary Not Exists]"
        Case Else
            Debug.Print "ErrNum[" & Err.Number & "]" & Err.Description
    End Select
    Debug.Print VarType(Var)
End Sub
'//////////////////////////////////////////////////////////////////////////
' �󂯎�����ϐ��ɂ��āA�^�t���[�ŃC�~�f�B�G�C�g�֏o�͂���
' @DP
Sub DP(ArgumentVariant As Variant)
        Dim BufferString As String
        Dim BufferVariant As Variant
        Dim i As Integer
    '========== Begin ==========
On Error GoTo ErrRtn
'    Debug.Print VarType(ArgumentVariant) '*
    '����`�ł�������z��́F8200
    'Array 8192
    'String 8
    If VarType(ArgumentVariant) = vbVariant Then Stop
    Select Case VarType(ArgumentVariant)
        Case vbEmpty: Debug.Print "vbEmpty"
        Case vbObject
            '�z��łȂ�
            If IsObject(ArgumentVariant) Then
                If IsError(ArgumentVariant) Then
                    BufferString = "Object.Name = " & ArgumentVariant.Name '�ϊ�
                Else
                    BufferString = "UnKnown Object" '�ϊ�
                End If
            Else
                BufferString = "UnKnown Object" '�ϊ�
            End If
            Debug.Print BufferString
        Case vbObject + vbArray
            If VarType(ArgumentVariant) < vbArray Then
                '�z��łȂ�
                If IsObject(ArgumentVariant) Then
                    If IsError(ArgumentVariant.Name) Then
                        BufferString = "UnKnown Object" '�ϊ�
                    Else
                        BufferString = "Object.Name = " & ArgumentVariant.Name '�ϊ�
                    End If
                Else
                    BufferString = "UnKnown Object" '�ϊ�
                End If
                Debug.Print BufferString
            Else
                '�z��
                'VarType(Var) >= vbArray
                For i = 0 To UBound(ArgumentVariant)
                    If IsObject(ArgumentVariant) Then
                        If IsError(ArgumentVariant) Then
                            BufferString = "UnKnown Object" '�ϊ�
                        Else
                            BufferString = "Object.Name = " & ArgumentVariant.Name '�ϊ�
                        End If
                    Else
                        BufferString = "UnKnown Object" '�ϊ�
                    End If
                    Debug.Print format(i, "000") & ":" & BufferString
                Next i
            End If
        Case Else
            If VarType(ArgumentVariant) < vbArray Then
                '�z��łȂ�
                BufferString = CnvVarToStr(ArgumentVariant) '�ϊ�
                Debug.Print BufferString
            Else
                '�z��
                'VarType(Var) >= vbArray
                For i = 0 To UBound(ArgumentVariant)
                    BufferVariant = ArgumentVariant(i)
                    BufferString = CnvVarToStr(BufferVariant) '�ϊ�
                    Debug.Print format(i, "000") & ":" & BufferString
                Next i
            End If
    End Select
Exit Sub
    '========== End =========
ErrRtn:
    Select Case Err.Number
        Case 9
            Debug.Print "[Ary Not Exists]"
        Case Else
            Debug.Print "ErrNum[" & Err.Number & "]" & Err.Description
    End Select
    Debug.Print VarType(Var)
End Sub
'//////////////////////////////////////////////////////////////////////////
Function ErrOut(argString As String)
    Debug.Print "" & Err.Number & " : " & Err.Description
    Debug.Print "Func.Name = " & argString
End Function
'//////////////////////////////////////////////////////////////////////////
Function DPErr()
    Debug.Print "Debug.Print Err : " & Err.Number & " : " & Err.Description
'    If Er.DebugMode = 1 Then
'        Stop
'    End If
End Function
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'@Str
'//////////////////////////////////////////////////////////////////////////
'�^�O����肫�� ���ׂ�
'() <> �Ȃǂň͂܂ꂽ���̂���肫�� StrCut���[�v��
Function StringCutTagAll( _
        ByVal BaseString As String, _
        BeforeString As String, _
        AfterString As String) _
        As String
    '===========================
    Dim BeforeInteger  As Integer
    Dim AfterInteger As Integer
    Dim CntErr As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    BeforeInteger = 0
    AfterInteger = 0
    CntErr = 0
    Do
        BeforeInteger = InStr(1, BaseString, AfterString, vbBinaryCompare)
        If BeforeInteger > 0 Then
            AfterInteger = _
                InStr(BeforeInteger, BaseString, AfterString, vbBinaryCompare _
            )
            If AfterInteger > 0 Then
                BaseString = StringCutTag(BaseString, BaseString, AfterString)
            End If
        End If
        CntErr = CntErr + 1
        If CntErr > Len(BaseString) Then Exit Do
    Loop While (BeforeInteger > 0 And AfterInteger > 0)
'        Debug.Print StrBase: Stop '*
    StringCutTagAll = BaseString
    Exit Function
    '========== End =========
ErrRtn:
    StringCutTagAll = BaseString
End Function
'//////////////////////////////////////////////////////////////////////////
'����StrTagCut�̒u��������
'�؂������Ƃ���ɕ�����}������@StrCutReplace 1�x�̂�
Function StringCutTagAndReplace( _
        ByVal BaseString As String, _
        BeforeString As String, _
        AfterString As String, _
        ReplaceString As String _
        ) As String
    '===========================
    Dim BeforeInteger As Integer
    Dim AfterInteger As Integer
    Dim TempInteger As Integer
    '========== Begin ==========
On Error GoTo ErrRtn
    BeforeInteger = InStr(1, BaseString, BeforeString, vbBinaryCompare)
    If nBefore > 0 Then
        AfterInteger = InStr(BeforeInteger, BaseString, AfterString, vbBinaryCompare)
        If AfterInteger > 0 Then
            StringCutTagAndReplace = _
                Left(BaseString, BeforeString - 1) & _
                ReplaceString & Right(BaseString, Len(BaseString) - AfterInteger)
        Else
            GoTo ErrRtn
        End If
    Else
        GoTo ErrRtn
    End If
    Exit Function
    '========== End =========
ErrRtn:
    StringCutTagAndReplace = BaseString
End Function
'//////////////////////////////////////////////////////////////////////////
'����StrTagCut�̒u��������
'�؂������Ƃ���ɕ�����}������@StrCutReplace���[�v�ԁ@���ׂ�
Function StringCutTagAndReplaceAll( _
        ByVal BaseString As String, _
        BeforeString As String, _
        AfterString As String, _
        ReplaceString As String _
        ) As String
    '===========================
    Dim BeforeInteger As Integer
    Dim AfterInteger As Integer
    Dim TempInteger As Integer
    Dim nBef  As Integer, nAft As Integer
    Dim ErrCnt As Long
    '========== Begin ==========
On Error GoTo ErrRtn
    BeforeInteger = 0: AfterInteger = 0
    '�u��������\��̂��̂̒��ɁA����������Ώۂ��������ꍇ���[�v���邽��
    If InStr(1, BaseString, BeforeString, vbBinaryCompare) Then
        If InStr(1, ReplaceString, AfterString, vbBinaryCompare) Then
            Stop
            StringCutTagAndReplaceAll = BaseString
            Exit Function
        End If
    End If
    ErrCnt = 0
    Do
        BeforeInteger = InStr(1, BaseString, AfterString, vbBinaryCompare)
        If BeforeInteger > 0 Then
            AfterInteger = InStr(BeforeInteger, BaseString, AfterString, vbBinaryCompare)
            If AfterInteger > 0 Then
                '�u������
                BaseString = _
                    StrCutTagReplace( _
                    BaseString, _
                    BeforeInteger, _
                    AfterInteger, _
                    ReplaceString _
                )
            End If
        End If
        ErrCnt = ErrCnt + 1
        If Len(BaseString) < ErrCnt Then GoTo ErrRtn
'        Debug.Print StrBase '*
    Loop While (BeforeInteger > 0 And AfterInteger > 0)
    StringCutTagAndReplaceAll = BaseString
    Exit Function
    '========== End =========
ErrRtn:
    StringCutTagAndReplaceAll = BaseString
End Function
'//////////////////////////////////////////////////////////////////////////
'"<<<"�̂Ƃ���"<<"��Replace�����"<<"���c��̂�h��
'�����Ȃ�܂�Replace�Ŏ�肫��
Function StringReplaceAll( _
        ByVal BaseString As String, _
        BeforeReplaceString As String, _
        AfterReplaceString As String _
        ) As String
    '===========================
    Dim Num As Integer
    Dim ErrCnt As Long
    '========== Begin ==========
On Error GoTo ErrRtn
    Num = InStr(1, BaseString, BeforeString, vbBinaryCompare)
    ErrCnt = 0
    Do
        Num = InStr(1, BaseString, BeforeString, vbBinaryCompare)
        BaseString = Replace(BaseString, BeforeReplaceString, AfterReplaceString)
        ErrCnt = ErrCnt + 1
        If Len(BaseString) < ErrCnt Then GoTo ErrRtn
    Loop While (Num > 0)
'    Debug.Print StrBase: Stop  '*
    StringReplaceAll = BaseString
    Exit Function
    '========== End =========
ErrRtn:
    StringReplaceAll = BaseString
End Function
'/////////////////////////////////////////////////////////////////////////////
'*WildCard��������
'"*a*b*c.txt"
Function InStrWild _
    (StartNum As Long, BaseStr As String, CompareStr As String) As Long
    Dim FlagLng As Long
    Dim CompreStrSplit() As String
    Dim n As Long
    Dim cnt As Long
    Dim BufStr As String
    Dim i As Integer
On Error GoTo ErrRtn
FlagLng = 0
    '========== Begin ==========
    If InStr(1, CompareStr, "*", vbBinaryCompare) > 0 Then
        CompreStrSplit = Split(CompareStr, "*")
        cnt = 0
        n = StartNum
        BufStr = Right(BaseStr, Len(BaseStr) - StartNum + 1)
        For i = 0 To UBound(CompreStrSplit)
            If Not CompreStrSplit(i) = "" Or (i = 0) Then
                n = InStr(n, BufStr, CompreStrSplit(i), vbBinaryCompare)
            Else
                n = -2
                Exit For
            End If
        Next i
    Else
        '* nothing
        n = InStr(StartNum, BaseStr, CompareStr, vbBinaryCompare)
    End If
    '========== End ==========

InStrWild = n
Exit Function
ErrRtn: InStrWild = SetErrFlag(FlagLng)
End Function
'/////////////////////////////////////////////////////////////////////////////
Function StringCount(BaseString As String, CountToString As String) As Long
    '===========================
    Dim FlagLong As Long
    Dim Count As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    Count = 0
    FlagLong = 1
    If CountToString = "" Then GoTo ErrRtn
    Do While (FlagLong > 0)
        FlagLong = InStr(FlagLong, BaseString, CountToString, vbBinaryCompare)
        If FlagLong > 0 Then
            Count = Count + 1
            FlagLong = FlagLong + 1
        Else
            Exit Do
        End If
    Loop
    '==========  End  ==========
StringCount = Count
Exit Function
ErrRtn:
StringCount = -1
End Function
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'���W���[���ǂݍ���
Sub ModuleExportVBComponents(Path As String)

Dim tObj As Object, ExportPath As String

'    ExportPath = ThisWorkbook.Path & "\export_" & Format(Now, "YYYYMMDDhhmm")
    'D:\MyFolder\VBA_Module
'    Debug.Print ThisWorkbook.Path
    ExportPath = Path
    
    If Dir(ExportPath, vbDirectory) = "" Then
        Call MkDir(ExportPath)
    End If

    For Each tObj In ThisWorkbook.VBProject.VBComponents
'        Debug.Print tObj.Type
'        Debug.Print tObj.Name
        Select Case tObj.Type
            Case STANDARD_MODULE
                tObj.Export ExportPath & "\" & tObj.Name & ".bas"
            Case CLASS_MODULE
                tObj.Export ExportPath & "\" & tObj.Name & ".cls"
            Case USER_FORM
                tObj.Export ExportPath & "\" & tObj.Name & ".frm"
        End Select
    Next
        
End Sub
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
' �ݒ�t�@�C���ɏ����Ă���O�����C�u������ǂݍ��݂܂��B
Public Sub ModuleImportVBComponents(FolderPath As String, FileName As String)
    Dim FilePath As String, FP As Integer, TStr As String, _
        ModulePath As String, ModuleName As String
    ' �S���W���[�����폜
'    clear_modules
    
    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox "" & FolderPath & "�����݂��܂���B"
    Else
        FilePath = AbsPath(FileName, FolderPath)    ' ��΃p�X�ɕϊ�
        If Dir(FilePath) = "" Then
            MsgBox "�O�����C�u������`" & FilePath & "�����݂��܂���B"
            Exit Sub
        End If
    End If
    
    ' �ǂݎ��
    FP = FreeFile
    Open FilePath For Input As #FP
    Do Until EOF(FP)
        ' �P�s����
        Line Input #FP, TStr
        If Len(TStr) > 0 Then
            ModuleName = Mid(TStr, 3, Len(TStr) - 6)
            ModulePath = AbsPath(TStr, FolderPath)
            If Dir(ModulePath) = "" Then
                ' �G���[
                MsgBox "���W���[��" & ModulePath & "�͑��݂��܂���B"
                Exit Do
            Else
'                Stop
                If Not ModuleName = "Common_ExportImport" Then
                    '���W���[���`�F�b�N
                    If ModuleExists(ModuleName) Then
                        '����΍폜���Ď�荞��
                        ModuleClear (ModuleName)
                        Call ModuleInclude(ModulePath, FolderPath)
                    Else
                        '�Ȃ���΂��̂܂܎�荞��
                        Call ModuleInclude(ModulePath, FolderPath)
                    End If
                End If
            End If
        End If
    Loop
    Close #FP

    ThisWorkbook.Save
    
End Sub
'//////////////////////////////////////////////////////////////////
'���̃��[�N�u�b�N�Ƀ��W���[�������邩
Function ModuleExists(ModuleName As String) As Boolean
    Dim tObj As Object
    ModuleExists = False
    For Each tObj In ThisWorkbook.VBProject.VBComponents
'        Debug.Print tObj.Name
        If tObj.Name = ModuleName Then
            ModuleExists = True
            Exit Function
        End If
    Next

End Function
'//////////////////////////////////////////////////////////////////
' ���郂�W���[�����O������ǂݍ��݂܂��B
' �p�X��.�Ŏn�܂�ꍇ�́C���΃p�X�Ɖ��߂���܂��B
Sub ModuleInclude(ByVal FilePath As String, FolderPath As String)
    ' ��΃p�X�ɕϊ�
    FilePath = AbsPath(FilePath, FolderPath)
    
    ' �W�����W���[���Ƃ��ēo�^
    ThisWorkbook.VBProject.VBComponents.Import FilePath
End Sub
'//////////////////////////////////////////////////////////////////
' ���W���[�����폜
 Sub ModuleClear(ModuleName As String)
    Dim ComponentObj As Object
    For Each ComponentObj In ThisWorkbook.VBProject.VBComponents
        If (ComponentObj.Type = 1 Or ComponentObj.Type = 2) _
            And ComponentObj.Name = ModuleName Then
            ' ���̕W�����W���[�����폜
'            Stop
            ThisWorkbook.VBProject.VBComponents.Remove ComponentObj
            Exit Sub
        End If
    Next ComponentObj
End Sub
'//////////////////////////////////////////////////////////////////
' �t�@�C���p�X���΃p�X�ɕϊ����܂��B
Function AbsPath(FilePath As String, FolderPath As String)

    ' ��΃p�X�ɕϊ�
    If Left(FilePath, 1) = "." Then
        FilePath = FolderPath & Mid(FilePath, 2, Len(FilePath) - 1)
    End If
    
    AbsPath = FilePath

End Function
'//////////////////////////////////////////////////////////////////



'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////////////////////
'NEW 2014.10.27
'What:=�����Ώۃf�[�^
'After:=�������J�n����Z��,'LookIn:= xlFormulas(����),xlValues(�l),xlComents(�R�����g)
'LookAt:= xlWhole(���ׂĈ�v)
'MatchCase := True(�召������ʂ���),
'MatchByte := True(�S���p��ʂ���),
'�w��V�[�g���當�����T���A�h���X��Ԃ��A���S��v �����E�A�と��
'SearchOrder:= xlByRows(�����),xlByColumns(�s����)
'SearchDirection := xlNext(�������F�K��l),xlPrevious(�t)
Function CellsFindWithOptionInAddress_RtnAddress( _
        FindSheetName As String, _
        FindAddress As String, _
        FindValue As String, _
        Lookat_xlWhole_xlPart As Integer, _
        SearchOrder_xlByRows_xlByColumns As Integer, _
        SearchDirection_xlNext_xlPrevious As Integer, _
        MatchCase_Bool As Boolean, _
        MatchByte_Bool As Boolean _
    ) As String
    Dim RtnAddress As String
    Dim FindObject As Object
    Dim TStr As String
    Dim FindAddress As String
On Error GoTo ErrRtn
    '========== Begin ==========
'    tstr = Sheets(FindSheetName).Range(FindStartAddress).Address
    
    If Not IsAddress(FindAddress) Then GoTo ErrRtn
    If Not IsSheetExists(FindSheetName) Then GoTo ErrRtn
    If Not (FindValue = "") Then GoTo ErrRtn
    
    If SearchOrder_xlByRows_xlByColumns = 0 Then
        Set FindObject = Sheets(FindSheetName).Range(FindAddress).find( _
            What:=FindValue, _
            AFTER:=Range(FindStartAddress), _
            LookIn:=xlValues, _
            lookat:=Lookat_xlWhole_xlPart, _
            SearchDirection:=SearchDirection_xlNext_xlPrevious, _
            MatchCase:=MatchCase_Bool, _
            MatchByte:=MatchByte_Bool _
        )
    Else
        Set FindObject = Sheets(FindSheetName).Range(FindAddress).find( _
            What:=FindValue, _
            AFTER:=Range(FindStartAddress), _
            LookIn:=xlValues, _
            lookat:=Lookat_xlWhole_xlPart, _
            SearchDirection:=SearchDirection_xlNext_xlPrevious, _
            searchorder:=SearchOrder_xlByRows_xlByColumns, _
            MatchCase:=MatchCase_Bool, _
            MatchByte:=MatchByte_Bool _
        )
    End If

    If FindObject Is Nothing Then
        RtnAddress = ""
    Else
        RtnAddress = FindObject.Address
    End If
    CellsFindWithOptionInAddress_RtnAddress = RtnAddress
    '========== End ==========
Exit Function
ErrRtn: ErrOut
CellsFindWithOptionInAddress_RtnAddress = ""
End Function
'/////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'What:=�����Ώۃf�[�^
'After:=�������J�n����Z��,'LookIn:= xlFormulas(����),xlValues(�l),xlComents(�R�����g)
'**�@LookAt:= xlWhole(���ׂĈ�v),xlPart  �ꕔ����v����Z��������
'**�@MatchCase := True(�召������ʂ���),
'**�@MatchByte := True(�S���p��ʂ���),
'�w��V�[�g���當�����T���A�h���X��Ԃ��A���S��v �����E�A�と��
'SearchOrder:= xlByRows(������̂�),xlByColumns(�s�����̂�)
'SearchDirection := xlNext(�������F�K��l),xlPrevious(�t)
'�Y������Z�������ׂĒT���āA�A�h���X���X�g�i������z��j��Ԃ�
Function CellsFindAllOfAddress_RtnAddressList( _
            SheetName As String, _
            FindAddress As String, _
            FindString As String, _
            xlOption As String) As String()
    Dim FindRange As Range
    Dim ForRange As Range
    Dim AddressListString() As String
    Dim Count As Long
    
    Dim ClTempArray As New Cl_Array
On Error GoTo ErrRtn
    Count = 0
    '========== Begin ==========
'    ReDim AddressListString(0)
    If Not IsAddress(FindAddress) Then GoTo ErrRtn
    If Not IsSheetExists(SheetName) Then GoTo ErrRtn
    Set FindRange = Sheets(SheetName).Range(FindAddress)
    
    'for each �� 1����StrComp����@���v������̂�z���
    For Each ForRange In FindRange
        Select Case xlOption
            Case "xlWhole" '���S��v
                If StrComp(ForRange.Value, FindString, vbBinaryCompare) = 0 Then
                    '�z��ɕt��
'                    AddressListString = ArrayStringgRedimAndAppendForString_RtnArrayString( _
'                        AddressListString, ForRange.Address _
'                    )
                    AddressListString = ClTempArray.GetValueDirect _
                        .StringAddValueLastElements_RtnString( _
                            AddressListString, ForRange.Address _
                    )
                End If
            Case "xlPart"  '������v
                If InStr(1, ForRange.Value, FindString, vbBinaryCompare) > 0 Then
                    '�z��ɕt��
'                    AddressListString = ArrayStringgRedimAndAppendForString_RtnArrayString( _
'                        AddressListString, ForRange.Address _
'                    )
                    AddressListString = ClTempArray.GetValueDirect _
                        .StringAddValueLastElements_RtnString( _
                            AddressListString, ForRange.Address _
                    )
                End If
                'debug
                Count = Count + 1
                If False Then
                
                End If
                If Count Mod 100000 = 0 Then Stop
                If ForRange.Row >= 108456 Then Stop
            Case "xlPartBefore"
            Case "xlPartAfter"
            Case "xlNormal", ""
                
            Case Else
        End Select
    Next
    '========== End ==========
    Set FindRange = Nothing
    Set ForRange = Nothing
Set ClTempArray = Nothing
CellsFindAllOfAddress_RtnAddressList = AddressListString
Exit Function
ErrRtn:
    CellsFindAllOfAddress_RtnAddressList = AddressListString
Set ClTempArray = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////////
'NewNew
'�Y������Z�������ׂĒT���āA�A�h���X���X�g�i������z��j��Ԃ�
'�g�p�֐��FAfterRtnAD  ArrayIsErasedStringOfBlankRtnStringArray
'���S��v
Function CellsFindAllOfSheet_RtnAddressList( _
        SheetName As String, FindStr As String, Lookat_xlWhole_xlPart As Integer) As String()
    Dim FlagLng As Long
    Dim ADFind As String
    Dim ADList() As String
    Dim cnt As Integer
    Dim LngBefore As Long, LngNow As Long
    Dim LngBeforeRow As Long, LngNowRow As Long
    Dim LngBeforeCol As Long, LngNowCol As Long
    
    Dim TempCell As New Cl_Cell
    Dim ClTempArray As New Cl_Array
On Error GoTo ErrRtn
    '========== Begin ==========
    cnt = 0
'    ADFind = CellsFindWithOption_RtnAddress( _
'        SheetName, _
'        "A1", _
'        FindStr, _
'        Lookat_xlWhole_xlPart, _
'        xlNext _
'    )
    
    ADFind = TempCell.GetValueDirect.FindWithOptionInAddress_RtnAddress( _
        SheetName, _
        Sheets(SheetName).Cells.Address, _
         "A1", _
        FindStr, _
        Lookat_xlWhole_xlPart, _
        0, _
        xlNext, _
        False, _
        False _
    )
    
    ReDim ADList(cnt)
    ADList(cnt) = ADFind
    Do While (Not ADFind = "")
        ADFind = Range(ADFind).Offset(0, 1).Address '���̃A�h���X�Z�b�g
'        ADFind = Range("A" & Range(ADFind).Column).Address   '���̃A�h���X�Z�b�g
'        ADFind = AfterRtnAD(SheetName, ADFind, FindStr)
'        ADFind = CellsFindWithOption_RtnAddress( _
'            SheetName, _
'            ADFind, _
'            FindStr, _
'            Lookat_xlWhole_xlPart, _
'            xlNext _
'        )
    ADFind = TempCell.GetValueDirect.FindWithOptionInAddress_RtnAddress( _
        SheetName, _
        Sheets(SheetName).Cells.Address, _
        ADFind, _
        FindStr, _
        Lookat_xlWhole_xlPart, _
        0, _
        xlNext, _
        False, _
        False _
    )
        If ADFind = "" Then
            '�݂���Ȃ����� ADFind=""�@�ɂȂ�̂Œ��Ӂ@�Ƃ肠������������
            ADFind = ADList(cnt)
'            Stop
        End If
        '��ׂ�@�Ō�܂ł�������߂�̂� B55 -> B77 -> B33
        'Row�l���ׂ�
        LngBefore = Range(ADList(cnt)).Row '1�O
        LngNow = Range(ADFind).Row '��
        If LngBefore >= LngNow Then
            'Row�������ꍇCol�l������ׂ�
            If LngBefore = LngNow Then
                LngBefore = Range(ADList(cnt)).Column '1�O
                LngNow = Range(ADFind).Column '��
                If LngBefore >= LngNow Then
                    '������������Δ�����
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Else
'            Stop
        End If

'        '��ׂ�@�Ō�܂ł�������߂�̂� B55 -> B77 -> B33
'        'Row��Col�����̕����������ꍇ�͔�����@�A�����ぃ�����O
'        LngBeforeRow = Range(ADList(Cnt)).Row '1�O
'        LngBeforeCol = Range(ADList(Cnt)).Column '1�O
'        LngNowRow = Range(ADFind).Row '��
'        LngNowCol = Range(ADFind).Column '��
'        If (LngBeforeRow >= LngNowRow) And (LngBeforeCol >= LngNowCol) Then
'            Exit Do
'        End If


        cnt = cnt + 1
        ReDim Preserve ADList(cnt)
        ADList(cnt) = ADFind
    Loop
'    Debug.Print UBound(ADList) '*
'    Call DP(ADList): Stop  '*
    ADList = ClTempArray.GetValueDirect _
        .ArrayStringIsErasedValueOfBlank_RtnStringArray(ADList)  '�󕶎��������
    
    '========== End ==========
CellsFindAllOfSheet_RtnAddressList = ADList
Set TempCell = Nothing
Set ClTempArray = Nothing
Exit Function
ErrRtn:
    CellsFindAllOfSheet_RtnAddressList = ADList
    DPErr
Set TempCell = Nothing
Set ClTempArray = Nothing
End Function
'/////////////////////////////////////////////////////////////////////////////
'Function CellsAddressGetRowsCount()
'
'End Function
'/////////////////////////////////////////////////////////////////////////////
'@CellsEnd
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃f�[�^�́A�㑤�ƍ������������āA�������v����΂��̃Z���ɒl����������
'�g�p�֐��FTableOfAddressIsGet_FindStringOfLeftTop_RtnAddress
'Function TableWrite_MatchTopAndLeftLabelName_RtnFlagLng( _
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃f�[�^�́A�㑤�ƍ������������āA�������v����΂��̒l��Ԃ�
'Base
'Function TableDoSomeByMode_ValueToMatchTopAndLeftLabelName_RtnString(
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���A�h���X�̒�����A�w�蕶�����T���A���̉��̒l�����o��
'Function TableGet_ValueOfUnderToMatchValue_RtnString(
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�f�[�^�i���o������ɂ�����́j�́A'�w�肵�����o���̒l�����ׂĔz��ŕԂ�
'�g�p�֐��FTableGet_ValueToMatchLeftAndTopLabel_RtnStringArray
'Function TableGet_ValueListToMatchLabel_RtnStringArray(
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃f�[�^�i�����̕\�A��Ɍ��o�����݂�j�̃f�[�^�����o��
'�e�[�u���A�h���X�̃f�[�^�����o��(�z���)�@'�X�^�[�g�A�h���X�͍���A
'�X�^�[�g�A�h���X�̃f�[�^(�����΂񍶂̃f�[�^)�͂��ׂĖ��܂��Ă���Ɖ���
'�����΂��̃f�[�^�����ׂĖ��܂��Ă���Ɖ��肷��(EndXl���邽��)
'SameValLeftLabel :
    'LabelName �̒l���AStrValue �Ɠ����Ƃ��́A���̃��x������z��ŕԂ�
'ValAll :
    'LabelName �̒l�����ׂĕԂ��@������z���
'Function TableGet_ValueToMatchLeftAndTopLabel_RtnStringArray _
        (SheetName As String, _
        FindAD As String, _
        StartStr As String, _
        LabelNameChkData As String, _
        LabelNameGetData As String, _
        CaptionPos_LabelTOP_LabelLEFT As String, _
        RtnVal_SameValLeftLabel_ValAll, _
        StrValue As String) As String()
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���̒�����A�l��T���A���v�������̂̂P���̒l�𓾂� ,���S��v
'Function TableGet_ValueOfUnerOfMatchValue_RtnString( _
            SheetName As String, _
            TableAddress As String, _
            FindString As String _
        ) As String
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃f�[�^�́A�㑤�ƍ������������āA�������v����΂��̒l��Ԃ�
'Function TableGet_ValueToMatchTopAndLeftLabelName_RtnString( _
        SheetName As String, _
        ADTable As String, _
        LabelNameTop As String, _
        LabelNameLeft As String) _
        As String
    '===========================
'//////////////////////////////////////////////////////////////////////////
'����̕������T���āA�e�[�u���A�h���X���擾���āA
'���Ə�̃��x�����Ƀ}�b�`�����Ƃ���̒l��ǂݍ���
'Function TableOfGetData_FindStringOfLeftTop_RtnAddress( _
        SheetName As String, _
        StartStr As String, _
        ADFind As String, _
        LabelNameTop As String, _
        LabelNameLeft As String, _
        MODE_AD_VAL As String) _
        As String
    '===========================
'//////////////////////////////////////////////////////////////////////////
'SameValLeftLabel :
'LabelName �̒l���AStrValue �Ɠ����Ƃ��́A�w��̃��x�����̒l�𕶎���ŕԂ��z��
'��������΃G���[
'Function TableGet_LabelNameToMatchValueOfOtherLabelName_RtnString _
        (SheetName As String, _
        FindAD As String, _
        StartStr As String, _
        CheckLabelName As String, _
        GetValLabelName As String, _
        CaptionPos_LabelTOP_LabelLEFT As String, _
        StrValue As String) As String
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃Z���ɁA�f�[�^��������
'�l�Ɋ֌W�Ȃ������l����������
'Function TableWrite_SameValueToSingleLabelOfAll_RtnFlagLong( _
        SheetName As String, _
        ADTable As String, _
        LabelNameWrite As String, _
        LABELMODE_LEFT_TOP As String, _
        WriteData As String _
        )
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃Z���̂ɁA�f�[�^��������
'�����̏㑤�̃��x���f�[�^���A�����̎w�肵���l�ƍ��v�����
'���̃A�h���X�ɏ�������
'�g�p�֐��FTableDoSomeByMode_MatchMultiValueAndMultiLabel_RtyStringArray
'�g�p�֐��FIsArrayExists
'Function TableWrite_ValueToMatchLabelInAnotherLabelValue_RtnFlagLong( _
        SheetName As String, _
        ADTable As String, _
        LabelNameTopListVar As Variant, _
        ValueListVar As Variant, _
        WriteLabelName As String, _
        LABELMODE_LEFT_TOP As String, _
        WriteData As String _
        )
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃Z���́A�f�[�^�𓾂�
'�����̏㑤�̃��x���f�[�^���A�����̎w�肵���l�ƍ��v�������A���삷��
'RTNMODE_AD_VAL
'Function TableDoSomeByMode_MatchMultiValueAndMultiLabel_RtyStringArray( _
        SheetName As String, _
        ADTable As String, _
        LabelNameTopListVar As Variant, _
        ValueListVar As Variant, _
        RtnLabelName As String, _
        LABELMODE_LEFT_TOP As String, _
        RTNMODE_AD_VAL As String _
        ) As String()
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�Z���A�h���X�́A���낢��ȃf�[�^�擾�@�A�h���X�֘A
'�g�p�֐��F�Ȃ�
'Function TableGet_VriousDataAboutAddress_RtnString( _
    SheetName As String, _
    ADTable As String, _
    MODE_BEGINAD_ENDAD_BEGINROW_BEGINCOL_ENDROW_ENDCOL_CAPTOP_CAPLEFT _
    ) As String
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�����̃��x���f�[�^�Ƃ����Е��̒P�ꃉ�x���f�[�^����l�𕶎���z��œ���
'Mode�͕����̂ق������C���ŁAOther�̓��C���̔��Α�
'LEFT <=> TOP
'Function TableGet_ValueListToMatchLabelNameList_RtnStringArray( _
            GetSheetName As String, _
            TableAddress As String, _
            LabelList() As String, _
            LabelOther As String, _
            MODE_LEFT_TOP As String _
        ) As String()
    '===========================
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'�󂯎���������̃��x�����̃A�h���X��Ԃ�
'Function TableGet_AddressToMatchLabelNameList_RtnStringArray( _
            GetSheetName As String, _
            TableAddress As String, _
            LabelList() As String, _
            MODE_LEFT_TOP As String _
        ) As String()
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�̃f�[�^�́A���x���A�h���X
'Function TableGet_AddressOfCaption_RtnAddress( _
        SheetName As String, _
        ADTable As String, _
        MODE_LEFT_TOP As String _
        ) As String
    '===========================
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�Z���̏I���s
'�g�p�֐��FTableGet_VriousDataAboutAddress_RtnString
Function TableGet_RowOfEndTable_RtnLong( _
    SheetName As String, _
    ADTable As String _
    ) As Long
    '===========================
    Dim FlagLng As Long
    Dim BufStr As String
    Dim RtnLng As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    BufStr = _
        TableGet_VriousDataAboutAddress_RtnString(SheetName, ADTable, "EndRow")
    If BufStr = "" Then GoTo ErrRtn
    RtnLng = CLng(BufStr)
    '==========  End  ==========
TableGet_RowOfEndTable_RtnLong = RtnLng
Exit Function
ErrRtn:
TableGet_RowOfEndTable_RtnLong = SetErrFlag(FlagLng)
End Function
'//////////////////////////////////////////////////////////////////////////
'�e�[�u���^�Z���̎n�߂̍s
'�g�p�֐��FTableGet_VriousDataAboutAddress_RtnString
Function TableGet_RowOfBeginTable_RtnLong( _
    SheetName As String, _
    ADTable As String _
    ) As Long
    '===========================
    Dim FlagLng As Long
    Dim BufStr As String
    Dim RtnLng As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    BufStr = _
        TableGet_VriousDataAboutAddress_RtnString(SheetName, ADTable, "BeginRow")
    If BufStr = "" Then GoTo ErrRtn
    RtnLng = CLng(BufStr)
    '==========  End  ==========
TableGet_RowOfBeginTable_RtnLong = RtnLng
Exit Function
ErrRtn:
TableGet_RowOfBeginTable_RtnLong = SetErrFlag(FlagLng)
End Function
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////


'@Is
'Is�֐���Boolean�߂�l
'/////////////////////////////////////////////////////////////////////////////
'�z�񂪃[���ł��邩�ǂ����@�[��False�@�P�ȏ�True
'Used
Function IsArrayExists(CheckArrayVariant As Variant) As Boolean
On Error GoTo ErrRtn
    With Block
        Dim i As Integer
        For i = 0 To UBound(CheckArrayVariant)
            Exit For
        Next i
    End With
    IsArrayExists = True
    Exit Function
ErrRtn:
    IsArrayExists = False
End Function
'/////////////////////////////////////////////////////////////////////////////
'�z�񂪃[���ł��邩�ǂ����@�[��-1�@�P�ȏ�True
Function IsArrayOfStringExists( _
        CheckArrayString() As String _
        ) As Boolean
    '===========================
    Dim i As Integer
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    With Block
        For i = 0 To UBound(CheckArrayString)
            Exit For
        Next i
    End With
    '========== End =========
IsArrayOfStringExists = True
Exit Function
ErrRtn:
'    en = Err.Number
'    es = Err.Description
    If Err.Number = 9 Then
    End If
    IsArrayOfStringExists = False
End Function
'/////////////////////////////////////////////////////////////////////////////
'�����z�񂪃[���ł��邩�ǂ����@�[��False�@�P�ȏ�True
Function IsArrayOfIntegerExists(CheckInteger() As Integer) As Boolean
On Error GoTo ErrRtn
    '========== Begin ==========
    With Block
        Dim i As Integer
        For i = 0 To UBound(CheckInteger)
            Exit For
        Next i
    End With
    IsArrayOfIntegerExists = True
Exit Function
    '========== End =========
ErrRtn:
    If Err.Number = 9 Then
    End If
    IsArrayOfIntegerExists = False
End Function
'/////////////////////////////////////////////////////////////////////////////
Function IsArrayOfLongExists(TLngA() As Long) As Boolean
    Dim i As Integer
On Error GoTo ErrRtn
    '========== Begin ==========
    For i = 0 To UBound(TLngA)
        Exit For
    Next i
    IsArrayOfLongExists = True
Exit Function
    '========== End =========
ErrRtn:
    If Err.Number = 9 Then
    End If
    IsArrayOfLongExists = False
End Function
'//////////////////////////////////////////////////////////////////////////
Function IsAddress(ByVal AD As String) As Boolean
On Error GoTo ErrorHandler
    If AD = "" Then
        IsAddress = False
'        Stop
    Else
        If Range(AD).Address = Range(AD).Address Then
        End If
        IsAddress = True
    End If
Exit Function
ErrorHandler:
    IsAddress = False
    Call DPErr: Stop
End Function
'//////////////////////////////////////////////////////////////////////////
'�����W���P��iRows=1,Colmun=1�j�Ȃ�@True
'����̃Z���̒l�𓾂�Ƃ��ɁA�A�h���X�������͈͂��w��ŃG���[�ɂȂ邱�Ƃ�����
Function IsAddressOfSingle(RngAD As String) As Boolean
    Dim AD As String
On Error GoTo ErrorHandler
    AD = RngAD
    If ActiveSheet.Range(AD).Rows = 1 And _
        ActiveSheet.Range(AD).Columns = 1 Then
        IsAddressOfSingle = True
    Else
        IsAddressOfSingle = False
    End If
'    Stop
Exit Function
ErrorHandler: 'Call DPErr:
'�����Z���I���̏ꍇ
'13 : �^����v���܂���B
IsAddressOfSingle = False
End Function
'//////////////////////////////////////////////////////////////////////////
'�����񂪋� �Ȃ�� True
Function IsStringOfBlank(Var As Variant) As Boolean
On Error GoTo ErrRtn
    If VarType(Var) = vbString Then
        If Var = "" Then
            IsStringOfBlank = False
        Else
            IsStringOfBlank = True  '��
        End If
    Else
        IsStringOfBlank = False '�^���Ⴄ
    End If
ErrRtn:
    IsStringOfBlank = False
End Function
'//////////////////////////////////////////////////////////////////////////
'�������[���������͋󕶎���ł����True
Function IsBlankOrZero(Var As Variant) As Boolean
    Dim Flag As Boolean
On Error GoTo ErrorHandler
'    Debug.Print VarType(var)
'    Debug.Print TypeName(var)
'    Stop
    Flag = False
    Select Case TypeName(Var)
        Case "String"
            If Var = "" Then
                Flag = True
            Else
                Flag = False
            End If
        Case "Integer"
            If Var = 0 Then
                IsBlankZero = True
            Else
                Flag = False
            End If
        Case Else
            Flag = True
    End Select
    IsBlankOrZero = Flag
Exit Function
ErrorHandler:
'    Call DPErr: Stop
    IsBlankOrZero = Flag
End Function
'//////////////////////////////////////////////////////////////////////////
'�������[���������͋󕶎���ł����True �ό�
'�g�p�֐��FIsBlankOrZero
Function IsBlankOrZeroForParamVariant(ParamArray Ary() As Variant) As Boolean
    Dim tV As Variant, Flag As Boolean
    Dim i As Integer
On Error GoTo ErrorHandler
'    If IsArrayExists(Ary) Then
        Flag = True
        For i = 0 To UBound(Ary)
            If IsBlankOrZero(Ary(i)) Then
                Flag = False
            End If
        Next i
        IsBlankZeroMulti = Flag
'    End If
Exit Function
ErrorHandler:
'    Call DPErr: Stop
    IsBlankOrZeroForParamVariant = Flag
End Function
'//////////////////////////////////////////////////////////////////////////
'�[���ȉ��Ȃ��
Function IsUnderZero(Var As Variant) As Boolean
    Dim Flag As Boolean
On Error GoTo ErrRtn
    If VarType(Var) < vbArray Then
        Flag = False
        Select Case VarType(Var)
            Case vbBoolean: If CLng(Var) <= 0 Then Flag = True
            Case vbByte:    If CLng(Var) <= 0 Then Flag = True
            Case vbCurrency: If CLng(Var) <= 0 Then Flag = True '�ʉ�
'            Case vbDataObject:      Flag = False
'            Case vbDecimal:         Flag = False '10�i��
'            Case vbDate:            Flag = False
'            Case vbEmpty:       var = "Data Is Empty Type"
'            Case vbError:       var = "Data Is Error Type"
            Case vbString:  If CLng(Var) <= 0 Then Flag = True
            Case vbLong:    If CLng(Var) <= 0 Then Flag = True
'            Case vbNull:        var = ""
'            Case vbObject:      var = "Data Is Object Type"
            Case vbInteger: If CLng(Var) <= 0 Then Flag = True
            Case vbSingle:  If CLng(Var) <= 0 Then Flag = True
            Case vbDouble:  If CLng(Var) <= 0 Then Flag = True
'            Case vbUserDefinedType:      Var = ""
            Case vbVariant: If CLng(Var) <= 0 Then Flag = True
            Case Else:      If CLng(Var) <= 0 Then Flag = True
        End Select
    Else 'Array
        Flag = False
    End If
    IsUnderZero = Flag
Exit Function
ErrRtn:
    DPErr
    IsUnderZero = False
End Function
'////////////////////////////////////////////////////////////////////////////
Function ObjectIsNothing( _
        ArgObjectVariant As Variant) As Boolean
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim TempString As String
    TempString = ArgObjectVariant.Name
    '==========  End  ==========
    ObjectIsNothing = True
Exit Function
ErrRtn:
    ObjectIsNothing = False
End Function
'////////////////////////////////////////////////////////////////////////////
'New
'Base
'�t�B�[���h�����邩�`�F�b�N�@���S��v
'Function IsFieldExistsByRS_RtnLong( _
'        RS As DAO.Recordset, _
'        FldName As String) As Long
'    '===========================
'    Dim FlagLng As Long
'    Dim FLD As DAO.Field
'On Error GoTo ErrRtn
'FlagLng = 0
'    '========== Begin ==========
'    If RS Is Nothing Then FlagLng = -10: GoTo ErrRtn
'    If FldName = "" Then FlagLng = -20: GoTo ErrRtn
'
'    For Each FLD In RS.Fields
'        If FLD.Name = FldName Then
'            FlagLng = 1
'            Exit For
'        End If
'    Next
'    '==========  End  ==========
'IsFieldExistsByRS_RtnLong = FlagLng
'Exit Function
'ErrRtn: IsFieldExistsByRS_RtnLong = SetErrFlag(FlagLng)
'End Function
'////////////////////////////////////////////////////////////////////////////
'New
'Base
'�t�B�[���h�����邩�`�F�b�N�@���S��v
'Function IsFieldExistsByDaoRs_RtnLong( _
'        RS As DAO.Recordset, _
'        FldName As String) As Long
'    '===========================
'    Dim FlagLng As Long
'    Dim FLD As DAO.Field
'On Error GoTo ErrRtn
'FlagLng = 0
'    '========== Begin ==========
'    If RS Is Nothing Then FlagLng = -10: GoTo ErrRtn
'    If FldName = "" Then FlagLng = -20: GoTo ErrRtn
'
'    For Each FLD In RS.Fields
'        If FLD.Name = FldName Then
'            FlagLng = 1
'            Exit For
'        End If
'    Next
'    '==========  End  ==========
'IsFieldExistsByDaoRs_RtnLong = FlagLng
'Exit Function
'ErrRtn: IsFieldExistsByDaoRs_RtnLong = Com.SetErrFlag(FlagLng)
'End Function
'//////////////////////////////////////////////////////////////////////////
'RS�����݂��邩
'Function IsRSByDAOExists(RS As DAO.Recordset) As Boolean
'    Dim Flag As Boolean
'On Error GoTo ErrorHandler
'    Flag = False
'    If RS Is Nothing Then
'        Flag = False
'    Else
'        Flag = True
'    End If
'    IsRSByDAOExists = Flag
'Exit Function
'ErrorHandler:
'    IsRSByDAOExists = Flag
'End Function
'////////////////////////////////////////////////////////////////////////////
'RS�����݂��邩
'Function IsRecordsetExists(RS As DAO.Recordset) As Boolean
'    Dim Flag As Boolean
'On Error GoTo ErrorHandler
'    Flag = False
'    If RS Is Nothing Then
'    Else
'        Flag = True
'    End If
'    IsRecordsetExists = Flag
'Exit Function
'ErrorHandler:
'    DPErr
'    IsRecordsetExists = Flag
'End Function
'////////////////////////////////////////////////////////////////////////////
'MDB�̒���TblName������΁@flag>0
'New Base
'Function IsTableOfMDBByDaoExists_RtnLng( _
'        MDB As DAO.Database, _
'        TblName As String _
'        ) As Long
'    '===========================
'    Dim FlagLng As Long
'    Dim TDF As DAO.TableDef
'On Error GoTo ErrRtn
'FlagLng = 0
'    '========== Begin ==========
'    For Each TDF In MDB.TableDefs
'        If Left(TDF.Name, 4) <> "MSys" Then
'            If TblName = TDF.Name Then
'                FlagLng = FlagLng + 1
'            End If
'        End If
'    Next TDF
'    '==========  End  ==========
'IsTableOfMDBByDaoExists_RtnLng = FlagLng
'Exit Function
'ErrRtn:
'    IsTableOfMDBByDaoExists_RtnLng = SetErrFlag(FlagLng)
'End Function
'//////////////////////////////////////////////////////////////////////////
'�V�[�g���݃`�F�b�N ThisBook
Function IsSheetExists(CheckSheetNameString As String) As Boolean
    Dim TempSheetObject As Worksheet
    Dim FlagBool As Boolean
    Dim BookNameString As String
On Error GoTo ErrRtn
    '========== Begin ==========
    FlagBool = False
    BookNameString = ThisWorkbook.Name
    For Each TempSheetObject In Workbooks(BookNameString).Sheets
        If TempSheetObject.Name = CheckSheetNameString Then
            FlagBool = True
            Exit For
        End If
    Next
    IsSheetExists = FlagBool
Exit Function
    '========== End =========
ErrRtn:
    IsSheetExists = FlagBool
    DPErr
End Function
'//////////////////////////////////////////////////////////////////////////
'�V�[�g���݃`�F�b�N
Function IsSheetExistsForOtherBook( _
        BookName As String, SheetNameForCheck As String _
        ) As Boolean
    Dim TempObj As Object, Flag As Boolean
On Error GoTo ErrorHandler
    Flag = False
    If BookName = "" Then BookName = ThisWorkbook.Name
    For Each TempObj In Workbooks(BookName).Sheets
        If TempObj.Name = SheetNameForCheck Then
            Flag = True
            Exit For
        End If
    Next
    IsSheetExistsForOtherBook = Flag
Exit Function
ErrorHandler:
'    Call DPErr: Stop
    IsSheetExistsForOtherBook = Flag
End Function
'//////////////////////////////////////////////////////////////////////////
Function IsFileDriveExists(DriveName As String) As Boolean
    Dim ObjFSO As Object
    Dim Flag As Boolean
On Error GoTo ErrorHandler
    Flag = False
    If Len(DriveName) = 1 Then
        ''DriveName(��)"C"�h���C�u�����݂��邩�ǂ������ׂ܂�
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        If ObjFSO.DriveExists(DriveName) Then
            Flag = True
'            MsgBox "E�h���C�u�����݂��܂�"
        Else
'            MsgBox "E�h���C�u�͑��݂��܂���"
        End If
    End If
    Set ObjFSO = Nothing
    IsFileDriveExists = Flag
Exit Function
ErrorHandler:
    Set ObjFSO = Nothing
    IsFileDriveExists = False
    Call DPErr: Stop
End Function
'//////////////////////////////////////////////////////////////////////////
Function IsFolderExists(FolderPath As String) As Boolean
    Dim FlagLng As Long
On Error GoTo ErrRtn
    FlagLng = 0
    If Dir(FolderPath, vbDirectory) <> "" Then
        IsFolderExists = True
    Else
        IsFolderExists = False
    End If
Exit Function
ErrRtn:
    IsFolderExists = False
End Function
'//////////////////////////////////////////////////////////////////////////
'�t�@�C�����݃`�F�b�N
Function IsFileExists(FilePath As String) As Boolean
    Dim Flag As Boolean
On Error GoTo ErrorHandler
    Flag = False
    If Dir(FilePath) <> "" Then
        Flag = True
    End If
    IsFileExists = Flag
Exit Function
ErrorHandler:
    IsFileExists = Flag
End Function
'//////////////////////////////////////////////////////////////////////////
'�t�@�C���ɏ�������
Function WriteFile(FilePath As String, data As String) As Boolean
    Dim Flag As Boolean
On Error GoTo ErrorHandler
    Flag = False
    Open FilePath For Append As #1
    Print #1, data
    Close #1
    WriteFile = Flag
Exit Function
ErrorHandler:
    WriteFile = Flag
End Function
'//////////////////////////////////////////////////////////////////////////
'�t�@�C����ǂݍ���
Function ReadFile(FilePath As String, data As String) As String()
    Dim data(0) As String
On Error GoTo ErrorHandler
    ' ==============
    Dim buf As String, n As Long
    Open FilePath For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
            ReDim Preserve data(n)
            data(n) = buf
            n = n + 1
        Loop
    Close #1
    ' ==============
    ReadFile = data
Exit Function
ErrorHandler:
    ReadFile = data
End Function

'########################################################################################
'CommonModule
'########################################################################################
Const SHOW_ERROR_MSG_BOX As Integer = 1
Const SHOW_ERROR_DEBUG_PRINT As Integer = 2

'�G���[�������ɏo�͂���
Public Sub DisplayError( _
    Optional FunctionName As String = "", _
    Optional ErrShowMode As Integer = 2, _
    Optional plErrNum As Long = 0, _
    Optional psErrDesc As String = "", _
    Optional psErrApl As String = "", _
    Optional psErrModule As String = "", _
    Optional psErrProc As String = "", _
    Optional pvErrNote As Variant)

    On Error Resume Next
    
    '�����ɉ����n���Ȃ��Ƃ��́A�G���[�I�u�W�F�N�g���甭�����Ă���G���[���擾���ĕ\������
    If plErrNum = 0 Then
        If Err.Number = 0 Then
            '�Ȃɂ��G���[���Ȃ��Ƃ��͏I������
            Exit Sub
        End If
        'Err.Source
        Call DisplayError(FunctionName, ErrShowMode, Err.Number, Err.Description, Erl, psErrModule, psErrProc, pvErrNote)
    Else
        '�����ɃG���[��񂪂���Ƃ��͈ȉ������s����
    End If

    '�G���[�Ɋւ�������擾����
    Dim sBuffer As String
    If IsMissing(pvErrNote) = False Then
        sBuffer = "�G���[���������܂����B" & vbCrLf & _
                    "�G���[�ԍ�" & vbTab & vbTab & ":" & Space(1) & CStr(plErrNum) & vbCrLf & _
                    "�G���[���e" & vbTab & vbTab & ":" & Space(1) & psErrDesc & vbCrLf & _
                    "�v���W�F�N�g��" & vbTab & vbTab & ":" & Space(1) & psErrApl & vbCrLf & _
                    "���W���[����" & vbTab & vbTab & ":" & Space(1) & psErrModule & vbCrLf & _
                    "�v���V�[�W����" & vbTab & vbTab & ":" & Space(1) & psErrProc & vbCrLf & _
                    "���l" & vbTab & vbTab & ":" & Space(1) & pvErrNote & vbCrLf
    Else
        sBuffer = "�G���[���������܂����B" & vbCrLf & _
                    "�G���[�ԍ�" & vbTab & vbTab & ":" & Space(1) & CStr(plErrNum) & vbCrLf & _
                    "�G���[���e" & vbTab & vbTab & ":" & Space(1) & psErrDesc & vbCrLf & _
                    "�v���W�F�N�g��" & vbTab & vbTab & ":" & Space(1) & psErrApl & vbCrLf & _
                    "���W���[����" & vbTab & vbTab & ":" & Space(1) & psErrModule & vbCrLf & _
                    "�v���V�[�W����" & vbTab & vbTab & ":" & Space(1) & psErrProc & vbCrLf
    End If
    
    '���[�h�ɂ���ďo�͂𕪂���
    If ErrShowMode = SHOW_ERROR_MSG_BOX Then
        '���b�Z�[�W�{�b�N�X�֏o�͂���
        Call MsgBox(sBuffer, vbCritical + vbOKOnly)
    ElseIf ErrShowMode = SHOW_ERROR_DEBUG_PRINT Then
        '�C�~�f�B�G�C�g�֏o�͂���
        Debug.Print sBuffer
    End If
        
End Sub



