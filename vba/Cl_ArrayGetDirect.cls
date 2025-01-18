VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cl_ArrayGetDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pErrorDiscriptionArrayString() As String
Private pErrorNumberArrayInteger() As Integer
Private pErrorFunctionNameArrayString() As String
Public pSystemErrorFlagInteger As Integer

'////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    pSystemErrorFlagInteger = 0
End Sub
'////////////////////////////////////////////////////////////////////////////
'Private Property Let ErrorDiscriptionArrayString(argString() As String)
'    pErrorDiscriptionArrayString = argString
'End Property
Public Property Get ErrorDiscriptionArrayString() As String()
    ErrorDiscriptionArrayString = pErrorDiscriptionArrayString
End Property
Public Property Get ErrorNumberArrayInteger() As Integer()
    ErrorNumberArrayInteger = pErrorNumberArrayInteger
End Property
Public Property Get ErrorFunctionNameArrayString() As String()
    ErrorFunctionNameArrayString = pErrorFunctionNameArrayString
End Property
Public Property Get SystemErrorFlagInteger() As Integer
    SystemErrorFlagInteger = pSystemErrorFlagInteger
End Property
'////////////////////////////////////////////////////////////////////////////
'他のクラスなどに移るとシステムエラーが消えてしまうのでメモする
'この関数は各クラスに配置しているので変更時注意
'基底クラスは別関数 ※注意
'Err.Desctiption,Err.Number、関数名　に分けておく
Sub SaveSystemError( _
        SystemErrorNumberInteger As Integer, _
        SystemErrorDescription As String, _
        Optional FunctionName As String _
        )
    If SystemErrorNumberInteger = 0 Then
        Exit Sub
    End If
    pSystemErrorFlagInteger = 1
    Call SaveSystemErrorByModuleForBaseClass( _
        pErrorDiscriptionArrayString, _
        SystemErrorDescription, _
        pErrorNumberArrayInteger, _
        SystemErrorNumberInteger, _
        pErrorFunctionNameArrayString, _
        FunctionName _
    )
End Sub
'////////////////////////////////////////////////////////////////////////////
'受け取った変数を配列に追加する
'Append＝付加する
'List:配列、AddInt:カウンタ
'使用関数：
Function IntegerAddValueLastElements_RtnString( _
            BaseArrayInteger() As Integer, _
            ByVal AddInteger As Integer _
        ) As Integer()
    '===========================
    Dim CntLong As Long
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    If IsArrayOfIntegerExists(BaseArrayInteger) Then
        CntLong = UBound(BaseArrayInteger)
        ReDim Preserve BaseArrayInteger(CntLong + 1)
        BaseArrayInteger(CntLong + 1) = AddInteger
    Else
        ReDim BaseArrayInteger(0)
        BaseArrayInteger(0) = AddInteger
    End If
    '========== End ==========
IntegerAddValueLastElements_RtnString = BaseArrayInteger
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "IntegerAddValueLastElements")
IntegerAddValueLastElements_RtnString = BaseArrayInteger
End Function
'/////////////////////////////////////////////////////////////////////////////
'受け取った変数を配列に追加する
'Append＝付加する
'List:配列、AddInt:カウンタ
'使用関数：
Function LongAddValueLastElements_RtnString( _
            BaseArrayLong() As Long, _
            ByVal AddLong As Long _
        ) As Long()
    '===========================
    Dim CntLong As Long
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    If IsArrayOfLongExists(BaseArrayLong) Then
        CntLong = UBound(BaseArrayLong)
        ReDim Preserve BaseArrayLong(CntLong + 1)
        BaseArrayLong(CntLong + 1) = AddLong
    Else
        ReDim BaseArrayLong(0)
        BaseArrayLong(0) = AddLong
    End If
    '========== End ==========
LongAddValueLastElements_RtnString = BaseArrayLong
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "LongAddValueLastElements")
LongAddValueLastElements_RtnString = BaseArrayLong
End Function
'/////////////////////////////////////////////////////////////////////////////
Function StringAddValueLastElements_RtnString( _
        BaseArrayString() As String, _
        ByVal AddString As String _
        ) As String()
    '===========================
    Dim CntLong As Long
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    If IsArrayOfStringExists(BaseArrayString) Then
        CntLong = UBound(BaseArrayString)
        ReDim Preserve BaseArrayString(CntLong + 1)
        BaseArrayString(CntLong + 1) = AddString
    Else
        ReDim BaseArrayString(0)
        BaseArrayString(0) = AddString
    End If
    '========== End ==========

StringAddValueLastElements_RtnString = BaseArrayString
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "StringAddValueLastElements")
StringAddValueLastElements_RtnString = BaseArrayString
End Function
'/////////////////////////////////////////////////////////////////////////////
'Long型配列の次の要素番号を得る
'要素数以上を指定すると、エラーになるためこの関数で緩和
Function LongNextElementsValue_RtnLong( _
            ArgArrayLong() As Long, _
            ArgNumber As Long _
        ) As Long
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    LongNextElementsValue_RtnLong = ArgArrayLong(ArgNumber + 1)
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "LongNextElementsValue")
LongNextElementsValue_RtnLong = 0
End Function
'/////////////////////////////////////////////////////////////////////////////
'String型配列の次の要素番号を得る
'要素数以上を指定すると、エラーになるためこの関数で緩和
Function StringNextElementsValue_RtnString( _
            ArgArrayString() As String, _
            ArgNumber As Long _
        ) As String
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    StringNextElementsValue_RtnString = ArgArrayString(ArgNumber + 1)
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "StringNextElementsValue")
StringNextElementsValue_RtnString = "__ERROR__"
End Function
'/////////////////////////////////////////////////////////////////////////////
'文字列配列の最後の文字列を得る
Function LastElement_RtnString(ArgArrayString() As String) As String
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    If Not IsArray(ArgArrayString) Then GoTo ErrRtn
    LastElement_RtnString = ArgArrayString(UBound(ArgArrayString))
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "LastElement")
LastElement_RtnString = "__ERROR__"
End Function
'////////////////////////////////////////////////////////////////////////////
'文字列配列の最初の要素に1つ追加する
Function ArrayStringAddFirstElement_RtnArrayString( _
            ByRef ArgArrayString() As String, _
            AddString As String _
            ) As String()
    '===========================
    Dim TempArrayString() As String
    Dim i As Integer
On Error GoTo ErrRtn
    '========== Begin ==========
    i = UBound(ArgArrayString)
    ReDim Preserve TempArrayString(UBound(ArgArrayString) + 1)
    TempArrayString(0) = AddString
    For i = 0 To UBound(ArgArrayString)
        TempArrayString(i + 1) = ArgArrayString(i)
    Next i
    '==========  End  ==========
ArrayStringAddFirstElement_RtnArrayString = TempArrayString
Erase TempArrayString
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ArrayStringAddFirstElement")
ArrayStringAddFirstElement_RtnArrayString = TempArrayString
Erase TempArrayString
End Function
'////////////////////////////////////////////////////////////////////////////
'文字列配列を1つの文字列に変換 要素間は改行する
Function ArrayStringConvertStringSeparatedLine_RtnString( _
            ByRef ArgArrayString() As String _
            ) As String
    '===========================
    Dim RtnString As String
    Dim i As Integer
On Error GoTo ErrRtn
    '========== Begin ==========
    For i = 0 To UBound(ArgArrayString)
        RtnString = RtnString & ArgArrayString(i)
        RtnString = RtnString & vbNewLine
    Next i
    RtnString = Left(RtnString, Len(RtnString) - 1)
    '==========  End  ==========
ArrayStringConvertStringSeparatedLine_RtnString = RtnString
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ArrayStringConvertStringSeparatedLine")
ArrayStringConvertStringSeparatedLine_RtnString = RtnString
End Function
'////////////////////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////////////////////
'Long
'New 2014.01.30
'Long配列の最小値を求める
Function ArrayLongIsGetMinimumValue_RtnLong(ArrayLong() As Long) As Long
    '===========================
    Dim i As Long
    Dim MaxLong As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    If IsArrayOfLongExists(ArrayLong) Then
        MaxLong = ArrayLong(0)
        For i = 1 To UBound(ArrayLong)
            If MaxLong > ArrayLong(i) Then 'ここだけ違う
                MaxLong = ArrayLong(i)
            End If
        Next i
    Else
        MaxLong = 0
        GoTo ErrRtn
    End If
    '==========  End  ==========
ArrayLongIsGetMinimumValue_RtnLong = MaxLong
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ArrayLongIsGetMinimumValue")
ArrayLongIsGetMinimumValue_RtnLong = MaxLong
End Function
'/////////////////////////////////////////////////////////////////////////////
Function ArrayLongIsGetMaximumValue_RtnLong(ArrayLong() As Long) As Long
    '===========================
    Dim i As Long
    Dim MaxLong As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    If IsArrayOfLongExists(ArrayLong) Then
        MaxLong = ArrayLong(0)
        For i = 1 To UBound(ArrayLong)
            If MaxLong < ArrayLong(i) Then 'ここだけ違う
                MaxLong = ArrayLong(i)
            End If
        Next i
    Else
        MaxLong = 0
        GoTo ErrRtn
    End If
    '==========  End  ==========
ArrayLongIsGetMaximumValue_RtnLong = MaxLong
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ArrayLongIsGetMaximumValue")
ArrayLongIsGetMaximumValue_RtnLong = MaxLong
End Function
'/////////////////////////////////////////////////////////////////////////////
'要検討
'可変個数の整数値を比べてすべて同じなら True
'フィールド登録時、サイズ、タイプ、名前を比べるときに使う
Function ArrayIntegerOfParamIsSameUbound_RtnBool( _
        ParamArray AryInt() As Variant _
        ) As Boolean
    Dim i As Integer, j As Integer, Flag As Boolean
On Error GoTo ErrorHandler
    If VarType(AryInt) >= vbArray Then
        Flag = True
        For i = 0 To UBound(AryInt) - 1
            For j = 1 To UBound(AryInt)
                If AryInt(i) = AryInt(j) Then
                
                Else
                    Flag = False    '一つでも違えば False
                End If
            Next j
        Next i
    Else
        '配列でない
        Flag = False
    End If
    ArrayIntegerOfParamIsSameUbound_RtnBool = Flag
Exit Function
ErrorHandler:
Call SaveSystemError(Err.Number, Err.Description, "ArrayIntegerOfParamIsSameUbound")
    ArrayIntegerOfParamIsSameUbound_RtnBool = Flag
End Function
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////

'2つの文字列配列を一つに　AのあとにB
'使用関数：IsArrayExists
'使用関数：ArrayStringgRedimAndAppendForString_RtnArrayString
Function ArrayStringCombineAndArrayString_RtnArrayString( _
    FrontArrayString() As String, BackArrayString() As String) As String()
    '===========================
    Dim RtnArrayString() As String
On Error GoTo ErrRtn
    '========== Begin ==========
    'ErrCheck
'    If Not IsArrayExists(FrontArrayString) Then GoTo ErrRtn
    If Not IsArrayExists(BackArrayString) Then GoTo ErrRtn
    RtnArrayString = FrontArrayString
        Dim i As Integer
        For i = 0 To UBound(BackArrayString)
            RtnArrayString = StringAddValueLastElements_RtnString( _
                RtnArrayString, BackArrayString(i))
        Next i
ArrayStringCombineAndArrayString_RtnArrayString = RtnArrayString
Exit Function
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ArrayStringCombineAndArrayString")
ArrayStringCombineAndArrayString_RtnArrayString = RtnArrayString
End Function

'/////////////////////////////////////////////////////////////////////////////
'空文字列””を配列から消す
'使用関数：IsArrayExists
Function ArrayStringIsErasedValueOfBlank_RtnStringArray(TStrA() As String) As String()
    Dim es As String, en As Integer
    Dim cnt As Integer, TStrA2() As String, i As Integer
On Error GoTo ErrorHandler
'    es = VarType(TStra(UBound(TStra)))
    If IsArrayExists(TStrA) Then
        cnt = 0
        For i = 0 To UBound(TStrA)
            If Not (TStrA(i) = "") Then
                ReDim Preserve TStrA2(cnt)
                TStrA2(cnt) = TStrA(i)
                cnt = cnt + 1
            End If
        Next i
        ArrayStringIsErasedValueOfBlank_RtnStringArray = TStrA2
    Else
        ArrayStringIsErasedValueOfBlank_RtnStringArray = TStrA
    End If
'    DPVarA (TStrA2)
'    Stop
    Exit Function
ErrorHandler:
'    Call DPErr(Err.Number, Err.Description): Stop
Call SaveSystemError(Err.Number, Err.Description, "ArrayStringIsErasedValueOfBlank")
    If Err.Number = 9 Then
        ArrayStringIsErasedValueOfBlank_RtnStringArray = TStrA
    End If
End Function
'/////////////////////////////////////////////////////////////////////////////
'Param Variant -> ArrayString　引数は文字列前提
Function ConvertVariantInParamArrayToArrayString_RtnArrayString( _
            ArgVariant() As Variant _
        ) As String()
        
End Function
'/////////////////////////////////////////////////////////////////////////////
'Variant -> ArrayString
Function ConvertVariantToArrayString_RtnArrayString( _
            ArgVariant As Variant _
        ) As String()
    '===========================
    Dim RtnArrayString() As String
    Dim BufStr As String
    Dim i As Integer
    '===========================
    Dim FlagLong As Long
On Error GoTo ErrRtn
FlagLong = 0
    '========== Begin ==========
    If VarType(ArgVariant) = vbArray + vbString Then
        'Aryは配列
        If IsArrayExists(ArgVariant) Then
            For i = 0 To UBound(ArgVariant)
            '配列の配列かもしれないので 2次元
                '引数は 3次元配列ということ
            'その場合は0番のみ "Ary(0):"をくわえて
                If VarType(ArgVariant(i)) > vbArray Then
                    BufStr = CnvVarToStr("Array(" & i & ")(0):" & ArgVariant(i)(0))
                    '配列に追加
                    RtnArrayString = _
                        StringAddValueLastElements_RtnString( _
                            RtnArrayString, BufStr _
                    )
                Else
                    '単体 引数は２D
                    BufStr = CnvVarToStr(ArgVariant(i))
                    '配列に追加
                    RtnArrayString = _
                        StringAddValueLastElements_RtnString( _
                            RtnArrayString, BufStr _
                    )
                End If
            Next i
        Else
            'variant が配列ではない
'            BufStr = CnvVarToStr(ArgVariant)
'            Stop
        End If
    Else
        'Aryは単体
        BufStr = CnvVarToStr(ArgVariant)
        ReDim RtnArrayString(0)
        RtnArrayString(0) = BufStr
    End If
    '==========  End  ==========
ConvertVariantToArrayString_RtnArrayString = RtnArrayString
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantToArrayString")
ConvertVariantToArrayString_RtnArrayString = RtnArrayString
End Function
'//////////////////////////////////////////////////////////////////////////
'2つの文字列配列を一つに　AのあとにB
'使用関数：IsArrayExists
'使用関数：ArrayStringgRedimAndAppendForString_RtnArrayString
'Function ArrayStringCombineAndArrayString_RtnArrayString( _
'    FrontArrayString() As String, BackArrayString() As String) As String()
'    '===========================
'    Dim RtnArrayString() As String
'On Error GoTo ErrRtn
'    '========== Begin ==========
'    'ErrCheck
'    If Not IsArrayExists(FrontArrayString) Then GoTo ErrRtn
'    If Not IsArrayExists(BackArrayString) Then GoTo ErrRtn
'    RtnArrayString = FrontArrayString
'    With Block
'        Dim i As Integer
'        For i = 0 To UBound(BackArrayString)
'            RtnArrayString = StringAddValueLastElements_RtnString( _
'                RtnArrayString, BackArrayString(i))
'        Next i
'    End With
'ArrayStringCombineAndArrayString_RtnArrayString = RtnArrayString
'Exit Function
'    '==========  End  ==========
'ErrRtn: ArrayStringCombineAndArrayString_RtnArrayString = RtnArrayString
'End Function
'//////////////////////////////////////////////////////////////////////////
Function CnvVarToStr(ArgumentVariant As Variant) As String
    Dim BufferString As String
    Dim BufferVariant As Variant
    Dim i As Integer
    '========== Begin ==========
    If VarType(ArgumentVariant) = vbVariant Then Stop
    If VarType(ArgumentVariant) < vbArray Then
        BufferVariant = ArgumentVariant
        Select Case VarType(BufferVariant)
            Case vbBoolean:     BufferString = CStr(BufferVariant)
            Case vbByte:        BufferString = CStr(BufferVariant)
            Case vbCurrency:    BufferString = CStr(BufferVariant)  '通貨
            Case vbDataObject:  BufferString = "Data Is DataObject Type"
            Case vbDecimal:     BufferString = CStr(BufferVariant) '10進数
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
    Else '配列
        'VarType(Var) >= vbArray
        For i = 0 To UBound(ArgumentVariant)
            BufferVariant = ArgumentVariant(i)
            Select Case VarType(ArgumentVariant)
                Case vbBoolean:     BufferString = CStr(BufferVariant)
                Case vbByte:        BufferString = CStr(BufferVariant)
                Case vbCurrency:    BufferString = CStr(BufferVariant)  '通貨
                Case vbDataObject:  BufferString = "Data Is DataObject Type"
                Case vbDecimal:     BufferString = CStr(BufferVariant) '10進数
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
                    '配列の配列の場合がある。深追いはしない
                    If VarType(BufferVariant) > vbArray Then
                        BufferString = "Data is Array of Array"
                    End If
            End Select
        Next i
    End If
    CnvVarToStr = BufferString
    '========== End ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "CnvVarToStr")
    CnvVarToStr = "System Error : " & Err.Number & _
        " [" & Err.Description & "]"
End Function
'/////////////////////////////////////////////////////////////////////////////
'最大個数を得る　エラー対策
Function GetUbound_RtnLong( _
        BaseArrayString() As String _
        ) As Long
    '===========================
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    GetUbound_RtnLong = UBound(BaseArrayString)
    '========== End ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "GetUbound")
GetUbound_RtnLong = 0
End Function

