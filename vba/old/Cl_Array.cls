VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cl_Array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const BEFORE = 0
Const AFTER = 1

'デフォルト値
Const USE_FLAG_OF_INTEGER = 1
Const USE_FLAG_OF_LONG = 1
Const USE_FLAG_OF_SINGLE = 1
Const USE_FLAG_OF_DOUBLE = 1
Const USE_FLAG_OF_CURRENCY = 0
Const USE_FLAG_OF_DATE = 1
Const USE_FLAG_OF_STRING = 1
Const USE_FLAG_OF_OBJECT = 1
Const USE_FLAG_OF_ERROR = 0
Const USE_FLAG_OF_BOOLEAN = 1
Const USE_FLAG_OF_VARIANT = 0
Const USE_FLAG_OF_DATAOBJECT = 0
Const USE_FLAG_OF_DECIMAL = 0
Const USE_FLAG_OF_BYTE = 0

Public GetValueDirect As New Cl_ArrayGetDirect

Private pMainArrayTypeInteger As Integer


Private pMainArrayInteger() As Integer
Private pMainArrayLong() As Long
Private pMainArraySingle() As Single
Private pMainArrayDouble() As Double
Private pMainArrayString() As String
Private pMainArrayBoolean() As Boolean


Private pMainArrayByte() As Byte    'XXX
'バイト型   0 から 255 までの符号なし 8 ビット (1 バイト) の整数です。
'バイナリ データに文字列型 (String) の変数を使用しないでください。
'これは、ANSI 形式と Unicode 形式の変換時にデータの内容が破損する場合があるためです。
Private pMainArrayDate() As Date          '日付型
Private pMainArrayCrrency() As Currency   '通貨型'XXX
Private pMainArrayVariant() As Variant
Private pMainArrayObject() As Object


Private pAbsolutePosition As Long
Private pReplayVariant As Variant 'Replay 中継
Private pEOA As Boolean 'End Of Array
Private pBOA As Boolean 'Beginning Of Array
Private pElementCount As Long '配列の個数

Private pCountIntegerWhenGetSameValue As Long
Private pMaxCountIntegerWhenGetSameValue As Long
Private pPositionLongWhenGetSameValue As Long
'Private pFlagWhenGetSameValueString As String

Private pValueWhenGetSameValueVariant As Variant


'エラー変数
Private pDebugMode As Integer
Private pErrorDescriptionArrayString() As String
Private pErrorNumberArrayInteger() As Integer
Private pErrorFunctionNameArrayString() As String
Private pSystemErrorFlagInteger As Integer

'////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////
Public Property Let debugMode(ArgFlagInteger As Integer)
    pDebugMode = ArgFlagInteger
End Property
Public Property Get debugMode() As Integer
    debugMode = pDebugMode
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let MaxCountIntegerWhenGetSameValue(ArgMaxCount As Long)
    pMaxCountIntegerWhenGetSameValue = ArgMaxCount
End Property
Public Property Get MaxCountIntegerWhenGetSameValue() As Long
    pMaxCountIntegerWhenGetSameValue = pEOA
End Property
'////////////////////////////////////////////////////////////////////////////
Property Let AbsolutePosition(ArgLong As Long)
    'absolutePositon　は　Moveなど のみで変更
    'Add Del　は影響するPosition　Last　などでなければ移動しない

    'マイナス値　@@@

    pAbsolutePosition = ArgLong
    'ゼロなら BOA
'    If pAbsolutePosition = 0 Then
'        pBOA = True
'    Else
'        pBOA = False
'    End If
'    'Maxなら EOA
'    If pElementCount = pAbsolutePosition Then
'        pEOA = True
'    Else
'        pEOA = False
'    End If
    Call UpdatePositionStatus
End Property

Property Get AbsolutePosition() As Long
    AbsolutePosition = pAbsolutePosition
End Property
'////////////////////////////////////////////////////////////////////////////
'End Of Array (boolean)
Public Property Get EOA() As Boolean
    EOA = pEOA
End Property
'////////////////////////////////////////////////////////////////////////////
'Beginning Of Array (boolean)
Public Property Get BOA() As Boolean
    EOA = pBOA
End Property
'////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    pSystemErrorFlagInteger = 0  '@@@
    pMainArrayTypeInteger = 0
    Call AllEraseMainArray
    pMaxCountIntegerWhenGetSameValue = 100 '@@@
End Sub
'////////////////////////////////////////////////////////////////////////////
'Private Property Let ErrorDiscriptionArrayString(argString() As String)
'    pErrorDescriptionArrayString = argString
'End Property
Public Property Get ErrorDiscriptionArrayString() As String()
    ErrorDiscriptionArrayString = pErrorDescriptionArrayString
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
Private Sub SaveSystemError( _
        SystemErrorNumberInteger As Integer, _
        SystemErrorDescription As String, _
        Optional FunctionName As String _
        )
    If SystemErrorNumberInteger = 0 Then
        Exit Sub
    End If
    pSystemErrorFlagInteger = 1
    Call SaveSystemErrorByModuleForBaseClass( _
        pErrorDescriptionArrayString, _
        SystemErrorDescription, _
        pErrorNumberArrayInteger, _
        SystemErrorNumberInteger, _
        pErrorFunctionNameArrayString, _
        FunctionName _
    )
    Err.Description = ""
End Sub
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetVbTypeOfMainArrayInteger() As Integer
    GetVbTypeOfMainArrayInteger = pMainArrayTypeInteger
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetVbTypeOfMainArrayString() As String
    GetVbTypeOfMainArrayString = ConvertTypeFlagIntegerToString(pMainArrayTypeInteger)
End Property
'********************************************************************
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArrayInteger() As Integer()
    GetDirectMainArrayInteger = pMainArrayInteger
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArrayLong() As Long()
    GetDirectMainArrayLong = pMainArrayLong
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArraySingle() As Single()
    GetDirectMainArraySingle = pMainArraySingle
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArrayDouble() As Double()
    GetDirectMainArrayDouble = pMainArrayDouble
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArrayString() As String()
    GetDirectMainArrayString = pMainArrayString
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArrayBoolean() As Boolean()
    GetDirectMainArrayBoolean = pMainArrayBoolean
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArrayDate() As Date()
    GetDirectMainArrayDate = pMainArrayDate
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Get GetDirectMainArrayObject() As Object()
    GetDirectMainArrayObject = pMainArrayObject
End Property
'////////////////////////////////////////////////////////////////////////////
''********************************************************************
Public Property Let SetDirectMainArrayInteger(ArgArrayInteger() As Integer)
    Call AllEraseMainArray
    pMainArrayInteger = ArgArrayInteger
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbInteger
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let SetDirectMainArrayLong(ArgArrayLong() As Long)
    Call AllEraseMainArray
    pMainArrayLong = ArgArrayLong
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbLong
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let SetDirectMainArraySingle(ArgArraySingle() As Single)
    Call AllEraseMainArray
    pMainArraySingle = ArgArraySingle
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbSingle
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let SetDirectMainArrayDouble(ArgArrayDouble() As Double)
    Call AllEraseMainArray
    pMainArrayDouble = ArgArrayDouble
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbDouble
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let SetDirectMainArrayString(ArgArrayString() As String)
    Call AllEraseMainArray
    pMainArrayString = ArgArrayString
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbString
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let SetDirectMainArrayBoolean(ArgArrayBoolean() As Boolean)
    Call AllEraseMainArray
    pMainArrayBoolean = ArgArrayBoolean
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbBoolean
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let SetDirectMainArrayObject(ArgArrayObject() As Object)
    Call AllEraseMainArray
    pMainArrayObject = ArgArrayObject
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbObject
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Public Property Let SetDirectMainArrayDate(ArgArrayDate() As Date)
    Call AllEraseMainArray
    pMainArrayDate = ArgArrayDate
    AbsolutePosition = 0
    pMainArrayTypeInteger = vbDate
    UpdateElementCount
    UpdatePositionStatus
End Property
'////////////////////////////////////////////////////////////////////////////
Private Sub AllEraseMainArray()
    Erase pMainArrayInteger
    Erase pMainArrayLong
    Erase pMainArraySingle
    Erase pMainArrayDouble
    Erase pMainArrayString
    Erase pMainArrayBoolean
    Erase pMainArrayDate
    Erase pMainArrayObject
    pMainArrayTypeInteger = 0
    pElementCount = 0
    AbsolutePosition = 0
    pCountIntegerWhenGetSameValue = 0
    pValueWhenGetSameValueVariant = Empty
'    pMaxCountIntegerWhenGetSameValue = 100 '@@@
    pDebugMode = 0  '@@@
End Sub
'////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////
Private Function ConvertTypeFlagIntegerToString(vbFlagInteger As Integer) As String
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim RtnString As String
    Select Case vbFlagInteger
            Case vbEmpty: RtnString = "vbEmpty"
            Case vbNull: RtnString = "vbNull"
            Case vbInteger: RtnString = "vbInteger"
            Case vbLong: RtnString = "vbLong"
            Case vbSingle: RtnString = "vbSingle"
            Case vbDouble: RtnString = "vbDouble"
            Case vbCurrency: RtnString = "vbCurrency"
            Case vbDate: RtnString = "vbDate"
            Case vbString: RtnString = "vbString"
            Case vbObject: RtnString = "vbObject"
            Case vbError: RtnString = "vbError"
            Case vbBoolean: RtnString = "vbBoolean"
            Case vbVariant: RtnString = "vbVariant"
            Case vbDataObject: RtnString = "vbDataObject"
            Case vbDecimal: RtnString = "vbDecimal"
            Case vbByte: RtnString = "vbByte"
            Case Is >= vbArray: RtnString = "vbArray + " & (vbFlagInteger - vbArray)
            Case Else: RtnString = "Case Else"
    End Select
    ConvertTypeFlagIntegerToString = RtnString
    '==========  End  ==========
ErrRtn:
Exit Function
Call SaveSystemError(Err.Number, Err.Description, "SetArray")
ConvertTypeFlagIntegerToString = "__ERROR__"
End Function
'////////////////////////////////////////////////////////////////////////////
Sub DeleteOneLineSystemErrorForThisClass()
    Call DeleteOneLineSystemErrorByModuleForBaseClass( _
        pErrorDescriptionArrayString, _
        pErrorNumberArrayInteger, _
        pErrorFunctionNameArrayString _
    )
End Sub
'////////////////////////////////////////////////////////////////////////////
'セットした後はエラーチェックすること
'セット時にはタイプを指定するvbString,vbIntegerなどなど
Sub SetArray(ArgArrayVariant As Variant, ArgVbTypeInteger As Integer)
    'Argの型を判定していて、それに合った型のプライベート変数へ
    
On Error GoTo ErrRtn
    Dim TempVariant As Variant
    '========== Begin ==========
    '新規でセットするのですべてクリア(Erase)
    Call AllEraseMainArray
    '変換して中継Variant型の変数に代入
    TempVariant = ConvertAndPutInReplayVariant(ArgArrayVariant)
    'エラー時に1行消す
    If pSystemErrorFlagInteger = 1 Then Call DeleteOneLineSystemErrorForThisClass
    'このVariantをArgvbTypeFlagInteger　任意の型に変換
    '↑の関数だけでは型判定を勝手にされてしまうため、意図しない型に入ることが予想される
    TempVariant = ConvertTypeOfVariantContents(TempVariant, ArgVbTypeInteger)
    If pSystemErrorFlagInteger = 1 Then
        'ここの関数だけ2行
        Call DeleteOneLineSystemErrorForThisClass
        Call DeleteOneLineSystemErrorForThisClass
    End If
'    DP (TempVariant): Stop
    '変換した配列をもったVariantをPrivateのそれぞれの型の配列に入れる
    Call SetArrayMainMethod(TempVariant)
'    DP (pMainArrayString): Stop
    
    Call UpdateElementCount
    '==========  End  ==========
Exit Sub
ErrRtn:
pMainArrayTypeInteger = 0
Call SaveSystemError(Err.Number, Err.Description, "SetArray")
End Sub
'////////////////////////////////////////////////////////////////////////////
Function GetElementCount() As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    GetElementCount = pElementCount
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "GetElementCount")
    GetElementCount = -1
End Function

'////////////////////////////////////////////////////////////////////////////
'XX
Sub SetArrayDesignateType(ArgVariant As Variant, vbTypeInteger As Integer) '指定
On Error GoTo ErrRtn
    Dim RtnVariant As Variant
    '========== Begin ==========
    Select Case vbTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                RtnVariant = pMainArrayInteger
            Case vbLong
                RtnVariant = pMainArrayLong
            Case vbSingle
                RtnVariant = pMainArraySingle
            Case vbDouble
                RtnVariant = pMainArrayDouble
            Case vbCurrency
                RtnVariant = pMainArrayCrrency
            Case vbDate
                RtnVariant = pMainArrayDate
            Case vbString
                RtnVariant = pMainArrayVariant
            Case vbObject
                RtnVariant = pMainArrayObject
            Case vbError
            Case vbBoolean
                RtnVariant = pMainArrayBoolean
            Case vbVariant
                RtnVariant = pMainArrayString
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    SetArrayDesignateType = RtnVariant
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "SetArrayDesignateType")
    SetArrayDesignateType = RtnVariant
End Sub
'////////////////////////////////////////////////////////////////////////////
'変換して中継Variant型の変数に代入
Function ConvertAndPutInReplayVariant( _
        ArgVariant As Variant _
        ) As Variant
    Dim Count As Integer
    Dim TypeFlagInteger As Integer
    Dim ErrorMessage As String
    
    Dim TempArrayInteger() As Integer
    Dim TempArrayLong() As Long
    Dim TempArraySingle() As Single
    Dim TempArrayDouble() As Double
'    Dim TempArrayCrrency() As Crrency
    Dim TempArrayDate() As Date
    Dim TempArrayString() As String
    Dim TempArrayBoolean() As Boolean
    Dim TempArrayObject() As Object
    
    Dim ReturnVariant As Variant
    '========== Begin ==========
    Count = 0
    TypeFlagInteger = 0
    Select Case VarType(ArgVariant)
            Case vbEmpty
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbEmpty]"
                Call SaveSystemError(-10, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbNull
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbNull]"
                Call SaveSystemError(-vbNull, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbInteger
                ReDim Preserve TempArrayInteger(Count)
                TempArrayInteger(Count) = CInt(ArgVariant)
                TypeFlagInteger = vbInteger
            Case vbLong
                ReDim Preserve TempArrayLong(Count)
                TempArrayLong(Count) = CLng(ArgVariant)
                TypeFlagInteger = vbLong
            Case vbSingle
                ReDim Preserve TempArraySingle(Count)
                TempArraySingle(Count) = CSng(ArgVariant)
                TypeFlagInteger = vbSingle
            Case vbDouble
                ReDim Preserve TempArrayDouble(Count)
                TempArrayDouble(Count) = CDbl(ArgVariant)
                TypeFlagInteger = vbDouble
            Case vbCurrency
'                ReDim Preserve TempArrayCrrency(Count)
'                TempArrayCrrency(Count) = CCur(ArgVariant)
'                TypeFlagInteger = vbCurrency
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbCurrency]"
                Call SaveSystemError(-vbCurrency, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbDate
                ReDim Preserve TempArrayDate(Count)
                TempArrayDate(Count) = CDate(ArgVariant)
                TypeFlagInteger = vbDate
            Case vbString
                ReDim Preserve TempArrayString(Count)
                TempArrayString(Count) = CStr(ArgVariant)
                TypeFlagInteger = vbString
            Case vbObject
                ReDim Preserve TempArrayObject(Count)
                TempArrayObject(Count) = ArgVariant
                TypeFlagInteger = vbObject
'                'エラー：セットできない
'                ErrorMessage = "ConvertArray Error"
'                ErrorMessage = ErrorMessage & "[vbObject]"
'                Call SaveSystemError(-vbObject, ErrorMessage, "ConvertAndPutInReplayVariant")
'                GoTo ErrRtn
            Case vbError
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbError]"
                Call SaveSystemError(-vbError, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbBoolean
                ReDim Preserve TempArrayBoolean(Count)
                TempArrayBoolean(Count) = CBool(ArgVariant)
                TypeFlagInteger = vbBoolean
            Case vbVariant
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbVariant]"
                Call SaveSystemError(-vbVariant, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
'                ReDim Preserve TempArrayVariant(Count)
'                TempArrayVariant(Count) = CVar(ArgVariant)
'                TypeFlagInteger = vbVariant
            Case vbDataObject
                'エラー：セットできない
                ErrorMessage = "SetArray Error"
                ErrorMessage = ErrorMessage & "[vbDataObject]"
                Call SaveSystemError(-vbDataObject, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbDecimal
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbDecimal]"
                Call SaveSystemError(-vbDecimal, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbByte
                'エラー：実装していません
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbByte]"
                Call SaveSystemError(-vbByte, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
                '====================================================
            Case Is >= vbArray
                'ここでセット
                Dim ForVar As Variant
                For Each ForVar In ArgVariant
                    Select Case (VarType(ArgVariant) - vbArray)
                        Case vbEmpty
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbEmpty]"
                            Call SaveSystemError(-10, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case vbNull
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbNull]"
                            Call SaveSystemError(-vbNull, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case vbInteger
                            ReDim Preserve TempArrayInteger(Count)
                            TempArrayInteger(Count) = CInt(ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbInteger
                        Case vbLong
                            ReDim Preserve TempArrayLong(Count)
                            TempArrayLong(Count) = CLng(ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbLong
                        Case vbSingle
                            ReDim Preserve TempArraySingle(Count)
                            TempArraySingle(Count) = CSng(ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbSingle
                        Case vbDouble
                            ReDim Preserve TempArrayDouble(Count)
                            TempArrayDouble(Count) = CDbl(ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbDouble
                        Case vbCurrency
'                            ReDim Preserve TempArrayCrrency(Count)
'                            TempArrayCrrency(Count) = CCur(ForVar)
'                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbCurrency
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbCurrency]"
                            Call SaveSystemError(-vbCurrency, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case vbDate
                            ReDim Preserve TempArrayDate(Count)
                            TempArrayDate(Count) = CDate(ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbDate
                        Case vbString
                            ReDim Preserve TempArrayString(Count)
                            TempArrayString(Count) = CStr(ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbString
                        Case vbObject
                            ReDim Preserve TempArrayObject(Count)
                            TempArrayObject(Count) = (ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbObject
'                            'エラー：セットできない
'                            ErrorMessage = "ConvertArray Error"
'                            ErrorMessage = ErrorMessage & "[vbArray + vbObject]"
'                            Call SaveSystemError(-vbObject, ErrorMessage, "ConvertAndPutInReplayVariant")
'                            GoTo ErrRtn
                        Case vbError
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbError]"
                            Call SaveSystemError(-vbError, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case vbBoolean
                            ReDim Preserve TempArrayBoolean(Count)
                            TempArrayBoolean(Count) = CBool(ForVar)
                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbBoolean
                        Case vbVariant
'                            ReDim Preserve TempArrayVariant(Count)
'                            TempArrayVariant(Count) = CVar(ForVar)
'                            If TypeFlagInteger = 0 Then TypeFlagInteger = vbVariant
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbVariant]"
                            Call SaveSystemError(-vbVariant, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case vbDataObject
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbDataObject]"
                            Call SaveSystemError(-vbDataObject, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case vbDecimal
                            'エラー：セットできない
                            ErrorMessage = "SetArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbDecimal]"
                            Call SaveSystemError(-vbDecimal, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case vbByte
                            'エラー：実装していません
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbByte]"
                            Call SaveSystemError(-vbByte, ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case Is >= vbArray
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + vbArray]"
                            Call SaveSystemError(-VarType(ArgVariant), ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                        Case Else
                            'エラー：セットできない
                            ErrorMessage = "ConvertArray Error"
                            ErrorMessage = ErrorMessage & "[vbArray + CaseElse]"
                            Call SaveSystemError(-VarType(ArgVariant), ErrorMessage, "ConvertAndPutInReplayVariant")
                            GoTo ErrRtn
                            'vbArray + NUMBER
                    End Select
                    Count = Count + 1
                Next
                '====================================================
            Case Else
                'エラー：セットできない
                ErrorMessage = "SetArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + CaseElse]"
                Call SaveSystemError(-vbArray, ErrorMessage, "SetArray")
                GoTo ErrRtn
        End Select
    '一度変換したものを ReturnVariant
    Select Case TypeFlagInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                ReturnVariant = TempArrayInteger
            Case vbLong
                ReturnVariant = TempArrayLong
            Case vbSingle
                ReturnVariant = TempArraySingle
            Case vbDouble
                ReturnVariant = TempArrayDouble
            Case vbCurrency
'                ReturnVariant = TempArrayCrrency
            Case vbDate
                ReturnVariant = TempArrayDate
            Case vbString
                ReturnVariant = TempArrayString
            Case vbObject
                ReturnVariant = TempArrayObject
            Case vbError
            Case vbBoolean
                ReturnVariant = TempArrayBoolean
            Case vbVariant
'                ReturnVariant = TempArrayVariant
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    ConvertAndPutInReplayVariant = ReturnVariant
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertAndPutInReplayVariant")
    ConvertAndPutInReplayVariant = ReturnVariant
End Function
'////////////////////////////////////////////////////////////////////////////
'成形したデータ(ReplayVariant)を'プライベート変数にセットする
Sub SetArrayMainMethod(ArgVariant As Variant)
On Error GoTo ErrRtn
    Dim Count As Integer
    Dim ErrorMessage As String
    '========== Begin ==========
    '全消し
    Call AllEraseMainArray
    
    Select Case VarType(ArgVariant)
        Case vbEmpty, vbNull
            If pDebugMode = 1 Then
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbEmpty or vbNull]"
                Call SaveSystemError(-10, ErrorMessage, "SetArrayMainMethod")
            End If
        Case Else
    End Select
    
    Count = 0
    Dim ForVar As Variant
    For Each ForVar In ArgVariant
        Select Case (VarType(ArgVariant) - vbArray)
                Case vbEmpty
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbEmpty]"
                    Call SaveSystemError(-10, ErrorMessage, "SetArrayMainMethod")
                Case vbNull
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbNull]"
                    Call SaveSystemError(-vbNull, ErrorMessage, "SetArrayMainMethod")
                Case vbInteger
                    ReDim Preserve pMainArrayInteger(Count)
                    pMainArrayInteger(Count) = CInt(ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbInteger
                Case vbLong
                    ReDim Preserve pMainArrayLong(Count)
                    pMainArrayLong(Count) = CLng(ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbLong
                Case vbSingle
                    ReDim Preserve pMainArraySingle(Count)
                    pMainArraySingle(Count) = CSng(ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbSingle
                Case vbDouble
                    ReDim Preserve pMainArrayDouble(Count)
                    pMainArrayDouble(Count) = CDbl(ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbDouble
                Case vbCurrency
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbCurrency]"
                    Call SaveSystemError(-vbCurrency, ErrorMessage, "SetArrayMainMethod")
'                    RtnVariant = pMainArrayCrrency
                Case vbDate
                    ReDim Preserve pMainArrayDate(Count)
                    pMainArrayDate(Count) = CDate(ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbDate
                Case vbString
                    ReDim Preserve pMainArrayString(Count)
                    pMainArrayString(Count) = CStr(ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbString
                Case vbObject
                    ReDim Preserve pMainArrayObject(Count)
                    Set pMainArrayObject(Count) = (ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbObject
'                    'エラー：セットできない
'                    ErrorMessage = "ConvertArray Error"
'                    ErrorMessage = ErrorMessage & "[vbArray + vbObject]"
'                    Call SaveSystemError(-vbObject, ErrorMessage, "SetArrayMainMethod")
                Case vbError
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbError]"
                    Call SaveSystemError(-vbError, ErrorMessage, "SetArrayMainMethod")
                Case vbBoolean
                    ReDim Preserve pMainArrayBoolean(Count)
                    pMainArrayBoolean(Count) = CBool(ForVar)
                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbBoolean
                Case vbVariant
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbVariant]"
                    Call SaveSystemError(-vbVariant, ErrorMessage, "SetArrayMainMethod")
'                    ReDim Preserve pMainArrayVariant(Count)
'                    pMainArrayVariant(Count) = CVar(ForVar)
'                    If pMainArrayTypeInteger = 0 Then pMainArrayTypeInteger = vbVariant
                Case vbDataObject
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbDataObject]"
                    Call SaveSystemError(-vbDataObject, ErrorMessage, "SetArrayMainMethod")
                Case vbDecimal
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbDecimal]"
                    Call SaveSystemError(-vbDecimal, ErrorMessage, "SetArrayMainMethod")
                Case vbByte
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbByte]"
                    Call SaveSystemError(-vbByte, ErrorMessage, "SetArrayMainMethod")
                Case Is >= vbArray
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[vbArray + vbArray]"
                    Call SaveSystemError(-vbArray, ErrorMessage, "SetArrayMainMethod")
                Case Else
                    'エラー：セットできない
                    ErrorMessage = "ConvertArray Error"
                    ErrorMessage = ErrorMessage & "[Case Else]"
                    Call SaveSystemError(-(VarType(ArgVariant) - vbArray), ErrorMessage, "SetArrayMainMethod")
        End Select
        Count = Count + 1
    Next
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "SetArrayMainMethod")
End Sub
'////////////////////////////////////////////////////////////////////////////
'成形されたデータ(中身が配列のVariant型変数)の中身の配列の型を変換する
Private Function ConvertTypeOfVariantContents( _
            ArgVariant As Variant, _
            ArgVbTypeInteger As Integer _
        ) As Variant
On Error GoTo ErrRtn
    Dim Count As Integer
    Dim ErrorMessage As String
    
    Dim TempArrayInteger() As Integer
    Dim TempArrayLong() As Long
    Dim TempArraySingle() As Single
    Dim TempArrayDouble() As Double
'    Dim TempArrayCrrency() As Crrency
    Dim TempArrayDate() As Date
    Dim TempArrayString() As String
    Dim TempArrayBoolean() As Boolean
    Dim TempArrayObject() As Object
    
    Dim ReturnVariant As Variant
    '========== Begin ==========
    Count = 0
    Dim ForVar As Variant
    For Each ForVar In ArgVariant
        Select Case ArgVbTypeInteger
                Case vbEmpty    '////Error
                    GoTo ErrRtn
                Case vbNull    '////Error
                    GoTo ErrRtn
                Case vbInteger
                    ReDim Preserve TempArrayInteger(Count)
                    If IsNumeric(ForVar) Then
                        TempArrayInteger(Count) = CInt(ForVar)
                    Else
                        TempArrayInteger(Count) = 0
                    End If
                Case vbLong
                    ReDim Preserve TempArrayLong(Count)
                    If IsNumeric(ForVar) Then
                        TempArrayLong(Count) = CLng(ForVar)
                    Else
                        TempArrayLong(Count) = 0
                    End If
                Case vbSingle
                    ReDim Preserve TempArraySingle(Count)
                    If IsNumeric(ForVar) Then
                        TempArraySingle(Count) = CSng(ForVar)
                    Else
                        TempArraySingle(Count) = 0
                    End If
                Case vbDouble
                    ReDim Preserve TempArrayDouble(Count)
                    If IsNumeric(ForVar) Then
                        TempArrayDouble(Count) = CDbl(ForVar)
                    Else
                        TempArrayDouble(Count) = 0
                    End If
                Case vbCurrency    '////Error
                    GoTo ErrRtn
                Case vbDate
                    ReDim Preserve TempArrayDate(Count)
                    If IsDate(ForVar) Then
                        TempArrayDate(Count) = CDate(ForVar)
                    Else
                        TempArrayDate(Count) = 0
                    End If
                Case vbString
                    ReDim Preserve TempArrayString(Count)
                    TempArrayString(Count) = CStr(ForVar)
                Case vbObject
                    ReDim Preserve TempArrayObject(Count)
                    If IsObject(ForVar) Then
                        TempArrayObject(Count) = (ForVar)
                    Else
                        Set TempArrayObject(Count) = Nothing
                    End If
                Case vbError    '////Error
                    GoTo ErrRtn
                Case vbBoolean
                    ReDim Preserve TempArrayBoolean(Count)
                    TempArrayBoolean(Count) = CBool(ForVar)
                Case vbVariant
                    '////Error
'                    ReDim Preserve TempArrayVariant(Count)
'                    TempArrayVariant(Count) = (ForVar)
                Case vbDataObject    '////Error
                    GoTo ErrRtn
                Case vbDecimal    '////Error
                    GoTo ErrRtn
                Case vbByte    '////Error
                    GoTo ErrRtn
                Case Is >= vbArray    '////Error
                    GoTo ErrRtn
                Case Else    '////Error
                    GoTo ErrRtn
        End Select
        Count = Count + 1
    Next
'    DP (TempArrayInteger): Stop
    
    '一度変換したものを ReturnVariant へ
    Select Case ArgVbTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                ReturnVariant = TempArrayInteger
            Case vbLong
                ReturnVariant = TempArrayLong
            Case vbSingle
                ReturnVariant = TempArraySingle
            Case vbDouble
                ReturnVariant = TempArrayDouble
            Case vbCurrency
'                ReturnVariant = TempArrayCrrency
            Case vbDate
                ReturnVariant = TempArrayDate
            Case vbString
                ReturnVariant = TempArrayString
            Case vbObject
                ReturnVariant = TempArrayObject
            Case vbError
            Case vbBoolean
                ReturnVariant = TempArrayBoolean
            Case vbVariant
'                ReturnVariant = TempArrayVariant
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
'    DP (ReturnVariant): Stop
    
    ConvertTypeOfVariantContents = ReturnVariant
    '==========  End  ==========
Exit Function
ErrRtn:
ErrorMessage = "ConvertTypeOfVariantContents Error"
'DPErr
Call SaveSystemError(-1000 - ArgVbTypeInteger, ErrorMessage, "ConvertTypeOfVariantContents")
Call SaveSystemError(Err.Number, Err.Description, "ConvertTypeOfVariantContents")
    ConvertTypeOfVariantContents = ReturnVariant
End Function
'////////////////////////////////////////////////////////////////////////////
'プライベートのMain配列に追加する メイン関数
Sub AddElementDsignatePosition( _
            AddValueVariant As Variant, _
            AddElementNumber As Long, _
            Before_or_After As String _
            )
On Error GoTo ErrRtn
    Dim ElementNumberLong As Long
    Dim Before_or_After_FlagInteger As Long
    Dim AddPlusNumber As Long
    
    
    Dim RtnVariant As Variant ' xx
    '========== Begin ==========
    Select Case Before_or_After
        Case "BEFORE", "Before", "before"
            Before_or_After_FlagInteger = BEFORE
            AddPlusNumber = BEFORE
        Case "AFTER", "After", "after"
            Before_or_After_FlagInteger = AFTER
            AddPlusNumber = AFTER
        Case Else
    End Select
        
    Select Case pMainArrayTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                pMainArrayInteger = AddElementArrayInteger( _
                    pMainArrayInteger, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbLong
                pMainArrayLong = AddElementArrayLong( _
                    pMainArrayLong, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbSingle
                pMainArraySingle = AddElementArraySingle( _
                    pMainArraySingle, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbDouble
                pMainArrayDouble = AddElementArrayDouble( _
                    pMainArrayDouble, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbCurrency
'                RtnVariant = pMainArrayCrrency
            Case vbDate
                pMainArrayDate = AddElementArrayDate( _
                    pMainArrayDate, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbString
                pMainArrayString = AddElementArrayString( _
                    pMainArrayString, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbObject
                pMainArrayObject = AddElementArrayObject( _
                    pMainArrayObject, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbError
            Case vbBoolean
                pMainArrayBoolean = AddElementArrayBoolean( _
                    pMainArrayBoolean, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbVariant
                pMainArrayVariant = AddElementArrayVariant( _
                    pMainArrayVariant, _
                    AddValueVariant, _
                    pElementCount, _
                    AddElementNumber, _
                    AddPlusNumber _
                )
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementDsignatePosition")
End Sub

'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを整数型Integer配列のVariantにする
Private Function ConvertVariantoArrayInteger(ArgVariant As Variant) As Integer()
On Error GoTo ErrRtn
    Dim RtnArrayInteger() As Integer
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArrayInteger(Count)
        RtnArrayInteger(Count) = CInt(ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArrayInteger = RtnArrayInteger
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArrayInteger")
    ConvertVariantoArrayInteger = RtnArrayInteger
End Function
'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを Long 型の配列のVariantにする
Private Function ConvertVariantoArrayLong(ArgVariant As Variant) As Long()
On Error GoTo ErrRtn
    Dim RtnArrayLong() As Long
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArrayLong(Count)
        RtnArrayLong(Count) = CLng(ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArrayLong = RtnArrayLong
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArrayLong")
    ConvertVariantoArrayLong = RtnArrayLong
End Function
'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを Single 型の配列のVariantにする
Private Function ConvertVariantoArraySingle(ArgVariant As Variant) As Single()
On Error GoTo ErrRtn
    Dim RtnArraySingle() As Single
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArraySingle(Count)
        RtnArraySingle(Count) = CSng(ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArraySingle = RtnArraySingle
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArraySingle")
    ConvertVariantoArraySingle = RtnArraySingle
End Function
'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを Double 型の配列のVariantにする
Private Function ConvertVariantoArrayDouble(ArgVariant As Variant) As Double()
On Error GoTo ErrRtn
    Dim RtnArrayDouble() As Double
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArrayDouble(Count)
        RtnArrayDouble(Count) = CDbl(ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArrayDouble = RtnArrayDouble
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArrayDouble")
    ConvertVariantoArrayDouble = RtnArrayDouble
End Function
'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを Date 型の配列のVariantにする
Private Function ConvertVariantoArrayDate(ArgVariant As Variant) As Date()
On Error GoTo ErrRtn
    Dim RtnArrayDate() As Date
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArrayDate(Count)
        RtnArrayDate(Count) = CDate(ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArrayDate = RtnArrayDate
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArrayDate")
    ConvertVariantoArrayDate = RtnArrayDate
End Function
'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを Object 型の配列のVariantにする
Private Function ConvertVariantoArrayObject(ArgVariant As Object) As Object()
On Error GoTo ErrRtn
    Dim RtnArrayObject() As Object
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArrayObject(Count)
        RtnArrayObject(Count) = (ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArrayObject = RtnArrayObject
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArrayObject")
    ConvertVariantoArrayObject = RtnArrayObject
End Function
'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを Boolean 型の配列のVariantにする
Private Function ConvertVariantoArrayBoolean(ArgVariant As Variant) As Boolean()
On Error GoTo ErrRtn
    Dim RtnArrayBoolean() As Boolean
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArrayBoolean(Count)
        RtnArrayBoolean(Count) = CVar(ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArrayBoolean = RtnArrayBoolean
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArrayBoolean")
    ConvertVariantoArrayBoolean = RtnArrayBoolean
End Function
'////////////////////////////////////////////////////////////////////////////
'受け取ったVariantを String 型の配列のVariantにする
Private Function ConvertVariantoArrayString(ArgVariant As Variant) As String()
On Error GoTo ErrRtn
    Dim RtnArrayString() As String
    '========== Begin ==========
    Dim ForVar As Variant
    Dim Count As Integer
     
    Count = 0
    For Each ForVar In ArgVariant
        ReDim Preserve RtnArrayString(Count)
        RtnArrayString(Count) = CStr(ForVar)
        Count = Count + 1
    Next
    ConvertVariantoArrayString = RtnArrayString
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertVariantoArrayString")
    ConvertVariantoArrayString = RtnArrayString
End Function
'////////////////////////////////////////////////////////////////////////////
Private Function AddElementArrayInteger( _
            BaseArrayInteger() As Integer, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Integer()

On Error GoTo ErrRtn
    Dim RtnArrayInteger() As Integer
    Dim AddArrayInteger() As Integer
    '========== Begin ==========
    '入力値：配列もカバーする
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArrayInteger　で変換する
    AddArrayInteger = ConvertVariantoArrayInteger(AddValueVariant)
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayInteger) Then
            '配列がある
'            ReDim Preserve BaseArrayInteger(BaseArrayMaxCount + 1)
            Dim i As Long
            Dim TempArrayInteger() As Integer
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayInteger) Then Exit For
                ReDim Preserve TempArrayInteger(i)
                TempArrayInteger(i) = BaseArrayInteger(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayInteger)
                ReDim Preserve TempArrayInteger(AddElementNumber + i)
                TempArrayInteger(AddElementNumber + i) = AddArrayInteger(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayInteger) Then Exit For
                ReDim Preserve TempArrayInteger(UBound(TempArrayInteger) + 1)
                TempArrayInteger(UBound(TempArrayInteger)) = BaseArrayInteger(i)
            Next i
            RtnArrayInteger = TempArrayInteger
        Else
            '配列がない->新規
            RtnArrayInteger = AddArrayInteger
            BaseArrayMaxCount = UBound(RtnArrayInteger)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArrayInteger")
        'Lastに追加
        For i = 0 To UBound(AddArrayInteger)
            ReDim Preserve TempArrayInteger(UBound(TempArrayInteger) + 1)
            TempArrayInteger(UBound(TempArrayInteger) + 1) = AddArrayInteger(i)
        Next i
        RtnArrayInteger = TempArrayInteger
    End If
    AddElementArrayInteger = RtnArrayInteger
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayInteger")
    AddElementArrayInteger = RtnArrayInteger
End Function
'////////////////////////////////////////////////////////////////////////////
'Long
'BaseArray のAddElementNumber 番目に 同じ型の配列(Variant型) を追加挿入する
  'AddPlusNumberは 0か1のみを受け付け
  '0ならAddElementNumber 番目、1ならAddElementNumber＋１番目に追加する
Private Function AddElementArrayLong( _
            BaseArrayLong() As Long, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Long()

On Error GoTo ErrRtn
    Dim RtnArrayLong() As Long
    Dim AddArrayLong() As Long
    '========== Begin ==========
    '入力値：配列もカバーする
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArray　Long　で変換する
    AddArrayLong = ConvertVariantoArrayLong(AddValueVariant)
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayLong) Then
            '配列がある
            Dim i As Long
            Dim TempArrayLong() As Long
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayLong) Then Exit For
                ReDim Preserve TempArrayLong(i)
                TempArrayLong(i) = BaseArrayLong(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayLong)
                ReDim Preserve TempArrayLong(AddElementNumber + i)
                TempArrayLong(AddElementNumber + i) = AddArrayLong(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayLong) Then Exit For
                ReDim Preserve TempArrayLong(UBound(TempArrayLong) + 1)
                TempArrayLong(UBound(TempArrayLong)) = BaseArrayLong(i)
            Next i
            RtnArrayLong = TempArrayLong
        Else
            '配列がない->新規
            RtnArrayLong = AddArrayLong
            BaseArrayMaxCount = UBound(RtnArrayLong)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArrayLong")
        'Lastに追加
        For i = 0 To UBound(AddArrayLong)
            ReDim Preserve TempArrayLong(UBound(TempArrayLong) + 1)
            TempArrayLong(UBound(TempArrayLong) + 1) = AddArrayLong(i)
        Next i
        RtnArrayLong = TempArrayLong
    End If
    AddElementArrayLong = RtnArrayLong
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayLong")
    AddElementArrayLong = RtnArrayLong
End Function
'////////////////////////////////////////////////////////////////////////////
'Single
'BaseArray のAddElementNumber 番目に 同じ型の配列(Variant型) を追加挿入する
  'AddPlusNumberは 0か1のみを受け付け
  '0ならAddElementNumber 番目、1ならAddElementNumber＋１番目に追加する
Private Function AddElementArraySingle( _
            BaseArraySingle() As Single, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Single()

On Error GoTo ErrRtn
    Dim RtnArraySingle() As Single
    Dim AddArraySingle() As Single
    '========== Begin ==========
    '入力値：配列もカバーする
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArray　Single　で変換する
    AddArraySingle = ConvertVariantoArraySingle(AddValueVariant)
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArraySingle) Then
            '配列がある
            Dim i As Long
            Dim TempArraySingle() As Single
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArraySingle) Then Exit For
                ReDim Preserve TempArraySingle(i)
                TempArraySingle(i) = BaseArraySingle(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArraySingle)
                ReDim Preserve TempArraySingle(AddElementNumber + i)
                TempArraySingle(AddElementNumber + i) = AddArraySingle(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArraySingle) Then Exit For
                ReDim Preserve TempArraySingle(UBound(TempArraySingle) + 1)
                TempArraySingle(UBound(TempArraySingle)) = BaseArraySingle(i)
            Next i
            RtnArraySingle = TempArraySingle
        Else
            '配列がない->新規
            RtnArraySingle = AddArraySingle
            BaseArrayMaxCount = UBound(RtnArraySingle)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArraySingle")
        'Lastに追加
        For i = 0 To UBound(AddArraySingle)
            ReDim Preserve TempArraySingle(UBound(TempArraySingle) + 1)
            TempArraySingle(UBound(TempArraySingle) + 1) = AddArraySingle(i)
        Next i
        RtnArraySingle = TempArraySingle
    End If
    AddElementArraySingle = RtnArraySingle
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArraySingle")
    AddElementArraySingle = RtnArraySingle
End Function
'////////////////////////////////////////////////////////////////////////////
'Double
'BaseArray のAddElementNumber 番目に 同じ型の配列(Variant型) を追加挿入する
  'AddPlusNumberは 0か1のみを受け付け
  '0ならAddElementNumber 番目、1ならAddElementNumber＋１番目に追加する
Private Function AddElementArrayDouble( _
            BaseArrayDouble() As Double, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Double()

On Error GoTo ErrRtn
    Dim RtnArrayDouble() As Double
    Dim AddArrayDouble() As Double
    '========== Begin ==========
    '入力値：配列もカバーする
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArray　Double　で変換する
    AddArrayDouble = ConvertVariantoArrayDouble(AddValueVariant)
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayDouble) Then
            '配列がある
            Dim i As Long
            Dim TempArrayDouble() As Double
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayDouble) Then Exit For
                ReDim Preserve TempArrayDouble(i)
                TempArrayDouble(i) = BaseArrayDouble(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayDouble)
                ReDim Preserve TempArrayDouble(AddElementNumber + i)
                TempArrayDouble(AddElementNumber + i) = AddArrayDouble(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayDouble) Then Exit For
                ReDim Preserve TempArrayDouble(UBound(TempArrayDouble) + 1)
                TempArrayDouble(UBound(TempArrayDouble)) = BaseArrayDouble(i)
            Next i
            RtnArrayDouble = TempArrayDouble
        Else
            '配列がない->新規
            RtnArrayDouble = AddArrayDouble
            BaseArrayMaxCount = UBound(RtnArrayDouble)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArraySingle")
        'Lastに追加
        For i = 0 To UBound(AddArrayDouble)
            ReDim Preserve TempArrayDouble(UBound(TempArrayDouble) + 1)
            TempArrayDouble(UBound(TempArrayDouble) + 1) = AddArrayDouble(i)
        Next i
        RtnArrayDouble = TempArrayDouble
    End If
    AddElementArrayDouble = RtnArrayDouble
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayDouble")
    AddElementArrayDouble = RtnArrayDouble
End Function
'////////////////////////////////////////////////////////////////////////////
'Date
'BaseArray のAddElementNumber 番目に 同じ型の配列(Variant型) を追加挿入する
  'AddPlusNumberは 0か1のみを受け付け
  '0ならAddElementNumber 番目、1ならAddElementNumber＋１番目に追加する
Private Function AddElementArrayDate( _
            BaseArrayDate() As Date, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Date()

On Error GoTo ErrRtn
    Dim RtnArrayDate() As Date
    Dim AddArrayDate() As Date
    '========== Begin ==========
    '入力値：配列もカバーする
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArray　Date　で変換する
    AddArrayDate = ConvertVariantoArrayDate(AddValueVariant)
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayDate) Then
            '配列がある
            Dim i As Long
            Dim TempArrayDate() As Date
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayDate) Then Exit For
                ReDim Preserve TempArrayDate(i)
                TempArrayDate(i) = BaseArrayDate(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayDate)
                ReDim Preserve TempArrayDate(AddElementNumber + i)
                TempArrayDate(AddElementNumber + i) = AddArrayDate(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayDate) Then Exit For
                ReDim Preserve TempArrayDate(UBound(TempArrayDate) + 1)
                TempArrayDate(UBound(TempArrayDate)) = BaseArrayDate(i)
            Next i
            RtnArrayDate = TempArrayDate
        Else
            '配列がない->新規
            RtnArrayDate = AddArrayDate
            BaseArrayMaxCount = UBound(RtnArrayDate)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArraySingle")
        'Lastに追加
        For i = 0 To UBound(AddArrayDate)
            ReDim Preserve TempArrayDate(UBound(TempArrayDate) + 1)
            TempArrayDate(UBound(TempArrayDate) + 1) = AddArrayDate(i)
        Next i
        RtnArrayDate = TempArrayDate
    End If
    AddElementArrayDate = RtnArrayDate
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayDate")
    AddElementArrayDate = RtnArrayDate
End Function
'////////////////////////////////////////////////////////////////////////////
'Variant
'BaseArray のAddElementNumber 番目に 同じ型の配列(Variant型) を追加挿入する
  'AddPlusNumberは 0か1のみを受け付け
  '0ならAddElementNumber 番目、1ならAddElementNumber＋１番目に追加する
Private Function AddElementArrayVariant( _
            BaseArrayVariant() As Variant, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Variant()

On Error GoTo ErrRtn
    Dim RtnArrayVariant() As Variant
    Dim AddArrayVariant() As Variant
    '========== Begin ==========
    '入力値：配列もカバーする
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArray　Variant　で変換する
    AddArrayVariant = ConvertVariantoArrayVariant(AddValueVariant)
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayVariant) Then
            '配列がある
            Dim i As Long
            Dim TempArrayVariant() As Variant
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayVariant) Then Exit For
                ReDim Preserve TempArrayVariant(i)
                TempArrayVariant(i) = BaseArrayVariant(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayVariant)
                ReDim Preserve TempArrayVariant(AddElementNumber + i)
                TempArrayVariant(AddElementNumber + i) = AddArrayVariant(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayVariant) Then Exit For
                ReDim Preserve TempArrayVariant(UBound(TempArrayVariant) + 1)
                TempArrayVariant(UBound(TempArrayVariant)) = BaseArrayVariant(i)
            Next i
            RtnArrayVariant = TempArrayVariant
        Else
            '配列がない->新規
            RtnArrayVariant = AddArrayVariant
            BaseArrayMaxCount = UBound(RtnArrayVariant)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArraySingle")
        'Lastに追加
        For i = 0 To UBound(AddArrayVariant)
            ReDim Preserve TempArrayVariant(UBound(TempArrayVariant) + 1)
            TempArrayVariant(UBound(TempArrayVariant) + 1) = AddArrayVariant(i)
        Next i
        RtnArrayVariant = TempArrayVariant
    End If
    AddElementArrayVariant = RtnArrayVariant
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayVariant")
    AddElementArrayVariant = RtnArrayVariant
End Function
'////////////////////////////////////////////////////////////////////////////
'Boolean
'BaseArray のAddElementNumber 番目に 同じ型の配列(Variant型) を追加挿入する
  'AddPlusNumberは 0か1のみを受け付け
  '0ならAddElementNumber 番目、1ならAddElementNumber＋１番目に追加する
Private Function AddElementArrayBoolean( _
            BaseArrayBoolean() As Boolean, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Boolean()

On Error GoTo ErrRtn
    Dim RtnArrayBoolean() As Boolean
    Dim AddArrayBoolean() As Boolean
    '========== Begin ==========
    '入力値：配列もカバーする
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArray　Boolean　で変換する
    AddArrayBoolean = ConvertVariantoArrayBoolean(AddValueBoolean)
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayBoolean) Then
            '配列がある
            Dim i As Long
            Dim TempArrayBoolean() As Boolean
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayBoolean) Then Exit For
                ReDim Preserve TempArrayBoolean(i)
                TempArrayBoolean(i) = BaseArrayBoolean(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayBoolean)
                ReDim Preserve TempArrayBoolean(AddElementNumber + i)
                TempArrayBoolean(AddElementNumber + i) = AddArrayBoolean(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayBoolean) Then Exit For
                ReDim Preserve TempArrayBoolean(UBound(TempArrayBoolean) + 1)
                TempArrayBoolean(UBound(TempArrayBoolean)) = BaseArrayBoolean(i)
            Next i
            RtnArrayBoolean = TempArrayBoolean
        Else
            '配列がない->新規
            RtnArrayBoolean = AddArrayBoolean
            BaseArrayMaxCount = UBound(RtnArrayBoolean)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArraySingle")
        'Lastに追加
        For i = 0 To UBound(AddArrayBoolean)
            ReDim Preserve TempArrayBoolean(UBound(TempArrayBoolean) + 1)
            TempArrayBoolean(UBound(TempArrayBoolean) + 1) = AddArrayBoolean(i)
        Next i
        RtnArrayBoolean = TempArrayBoolean
    End If
    AddElementArrayBoolean = RtnArrayBoolean
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayBoolean")
    AddElementArrayBoolean = RtnArrayBoolean
End Function
'////////////////////////////////////////////////////////////////////////////
Private Function AddElementArrayString( _
            BaseArrayString() As String, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As String()

On Error GoTo ErrRtn
    Dim RtnArrayString() As String
    Dim AddArrayString() As String
    '========== Begin ==========
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArrayInteger　で変換する
    AddArrayString = ConvertVariantoArrayString(AddValueVariant)
'    DP (AddArrayString): Stop
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayString) Then
            '配列がある
'            ReDim Preserve BaseArrayString(BaseArrayMaxCount + 1)
            Dim i As Long
            Dim TempArrayString() As String
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayString) Then Exit For
                ReDim Preserve TempArrayString(i)
                TempArrayString(i) = BaseArrayString(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayString)
                ReDim Preserve TempArrayString(AddElementNumber + i)
                TempArrayString(AddElementNumber + i) = AddArrayString(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayString) Then Exit For
                ReDim Preserve TempArrayString(UBound(TempArrayString) + 1)
                TempArrayString(UBound(TempArrayString)) = BaseArrayString(i)
            Next i
            RtnArrayString = TempArrayString
        Else
            '配列がない->新規
            RtnArrayString = AddArrayString
            BaseArrayMaxCount = UBound(RtnArrayString)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArrayString")
        'Lastに追加
        'Lastに追加
        For i = 0 To UBound(AddArrayString)
            ReDim Preserve TempArrayString(UBound(TempArrayString) + 1)
            TempArrayString(UBound(TempArrayString) + 1) = AddArrayString(i)
        Next i
        RtnArrayString = TempArrayString
    End If
    AddElementArrayString = RtnArrayString
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayString")
    AddElementArrayString = RtnArrayString
End Function
'////////////////////////////////////////////////////////////////////////////
Private Function AddElementArrayObject( _
            BaseArrayObject() As Object, _
            AddValueVariant As Variant, _
            BaseArrayMaxCount As Long, _
            AddElementNumber As Long, _
            AddPlusNumber As Long _
        ) As Object()

On Error GoTo ErrRtn
    Dim RtnArrayObject() As Object
    Dim AddArrayObject() As Object
    '========== Begin ==========
    'AddValueVariantの中身は整数型配列
    'ConvertVariantToArray Object　で変換する @@@
    AddArrayObject = ConvertVariantoArrayObject(AddValueVariant)
'    DP (AddArrayObject): Stop
    
    If (AddElementNumber >= 0) And (AddElementNumber <= BaseArrayMaxCount) Then
        If IsArrayExists(BaseArrayObject) Then
            '配列がある
'            ReDim Preserve BaseArrayString(BaseArrayMaxCount + 1)
            Dim i As Long
            Dim TempArrayObject() As Object
            'Beforeなら0、Afterなら1加算
            AddElementNumber = AddElementNumber + AddPlusNumber
            'Tempにコピー 前半
            For i = 0 To AddElementNumber
                If i > UBound(BaseArrayObject) Then Exit For
                ReDim Preserve TempArrayObject(i)
                TempArrayObject(i) = BaseArrayObject(i)
            Next i
            'Tempにコピー中間
            For i = 0 To UBound(AddArrayObject)
                ReDim Preserve TempArrayObject(AddElementNumber + i)
                TempArrayObject(AddElementNumber + i) = AddArrayObject(i)
            Next i
            'Tempにコピー 後半
            For i = AddElementNumber To BaseArrayMaxCount
                If i > UBound(BaseArrayObject) Then Exit For
                ReDim Preserve TempArrayObject(UBound(TempArrayObject) + 1)
                TempArrayObject(UBound(TempArrayObject)) = BaseArrayObject(i)
            Next i
            RtnArrayObject = TempArrayObject
        Else
            '配列がない->新規
            RtnArrayObject = AddArrayObject
            BaseArrayMaxCount = UBound(RtnArrayObject)
        End If
    Else
        'ElementNumberが配列の番号(0～ElementCount)から外れている
        'Errorを出してLastに追加
        Dim ErrorMessage As String
        ErrorMessage = "AddElementNumberが範囲外です。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementArrayObject")
        'Lastに追加
        'Lastに追加
        For i = 0 To UBound(AddArrayObject)
            ReDim Preserve TempArrayObject(UBound(TempArrayObject) + 1)
            TempArrayObject(UBound(TempArrayObject) + 1) = AddArrayObject(i)
        Next i
        RtnArrayObject = TempArrayObject
    End If
    AddElementArrayObject = RtnArrayObject
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementArrayObject")
    AddElementArrayObject = RtnArrayObject
End Function
'////////////////////////////////////////////////////////////////////////////
Private Sub UpdatePositionStatus()
On Error GoTo ErrRtn
    Dim MaxCountLong As Long
    '========== Begin ==========
    'ゼロ以下はゼロ
    If pAbsolutePosition < -1 Then
        pAbsolutePosition = 0
    End If
    '削除されたときなど
    If pAbsolutePosition > pElementCount + 2 Then
        pAbsolutePosition = pElementCount + 1
    End If
    
    'Last + MoveNext => EOA
    If pAbsolutePosition = pElementCount + 1 Then
        pEOA = True
    Else
        pEOA = False
    End If
    
    'First + MovePrev => BOA
    If pAbsolutePosition = -1 Then
        pBOA = True
    Else
        pBOA = False
    End If
    
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "UpdatePositionStatus")
    pElementCount = 0
    If pElementCount = 0 Then pEOA = True
End Sub
'////////////////////////////////////////////////////////////////////////////
'メイン配列を Variantで取得　中身は変わらない
'MainArrayType型の配列が入ったVariantを取得
Function GetMainArrayAfterConversionToVariant() As Variant
On Error GoTo ErrRtn
    Dim RtnVariant As Variant
    '========== Begin ==========
    Select Case pMainArrayTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                RtnVariant = pMainArrayInteger
            Case vbLong
                RtnVariant = pMainArrayLong
            Case vbSingle
                RtnVariant = pMainArraySingle
            Case vbDouble
                RtnVariant = pMainArrayDouble
            Case vbCurrency
                RtnVariant = pMainArrayCrrency
            Case vbDate
                RtnVariant = pMainArrayDate
            Case vbString
                RtnVariant = pMainArrayString
            Case vbObject
                RtnVariant = pMainArrayObject
            Case vbError
            Case vbBoolean
                RtnVariant = pMainArrayBoolean
            Case vbVariant
                RtnVariant = pMainArrayVariant
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    GetMainArrayAfterConversionToVariant = RtnVariant
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "GetMainArrayAfterConversionToVariant")
    GetMainArrayAfterConversionToVariant = RtnVariant
End Function
'////////////////////////////////////////////////////////////////////////////
'メイン配列を Variantで取得　中身は変わらない
'MainArrayType型の配列が入ったVariantを取得
Function GetSingleElementAfterConversionToVariant() As Variant
On Error GoTo ErrRtn
    Dim RtnVariant As Variant
    '========== Begin ==========
    Select Case pMainArrayTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                RtnVariant = pMainArrayInteger(pAbsolutePosition)
            Case vbLong
                RtnVariant = pMainArrayLong(pAbsolutePosition)
            Case vbSingle
                RtnVariant = pMainArraySingle(pAbsolutePosition)
            Case vbDouble
                RtnVariant = pMainArrayDouble(pAbsolutePosition)
            Case vbCurrency
                RtnVariant = pMainArrayCrrency(pAbsolutePosition)
            Case vbDate
                RtnVariant = pMainArrayDate(pAbsolutePosition)
            Case vbString
                RtnVariant = pMainArrayString(pAbsolutePosition)
            Case vbObject
                Set RtnVariant = pMainArrayObject(pAbsolutePosition)
            Case vbError
            Case vbBoolean
                RtnVariant = pMainArrayBoolean(pAbsolutePosition)
            Case vbVariant
                RtnVariant = pMainArrayVariant(pAbsolutePosition)
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    If pMainArrayTypeInteger = 9 Then
        Set GetSingleElementAfterConversionToVariant = RtnVariant
        Call CheckValueOrTheSameWhenGetSingleElement
    Else
        GetSingleElementAfterConversionToVariant = RtnVariant
        Call CheckValueOrTheSameWhenGetSingleElement
    End If
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "GetSingleElementAfterConversionToVariant")
    If pMainArrayTypeInteger = 9 Then
        Set GetSingleElementAfterConversionToVariant = RtnVariant
        Call CheckValueOrTheSameWhenGetSingleElement
    Else
        GetSingleElementAfterConversionToVariant = RtnVariant
        Call CheckValueOrTheSameWhenGetSingleElement
    End If
End Function
'////////////////////////////////////////////////////////////////////////////
Private Sub UpdateElementCount()
On Error GoTo ErrRtn
    Dim MaxCountLong As Long
    '========== Begin ==========
    Select Case pMainArrayTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                MaxCountLong = UBound(pMainArrayInteger)
            Case vbLong
                MaxCountLong = UBound(pMainArrayLong)
            Case vbSingle
                MaxCountLong = UBound(pMainArraySingle)
            Case vbDouble
                MaxCountLong = UBound(pMainArrayDouble)
            Case vbCurrency
                MaxCountLong = UBound(pMainArrayCrrency)
            Case vbDate
                MaxCountLong = UBound(pMainArrayDate)
            Case vbString
                MaxCountLong = UBound(pMainArrayString)
            Case vbObject
                MaxCountLong = UBound(pMainArrayObject)
            Case vbError
            Case vbBoolean
                MaxCountLong = UBound(pMainArrayBoolean)
            Case vbVariant
'                MaxCountLong = UBound(pMainArrayVariant)
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    pElementCount = MaxCountLong
'    If MaxCountLong = 0 Then pEOA = True
    If IsArrayExists(GetMainArrayAfterConversionToVariant) Then
        pEOA = False
    End If
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "UpdateElementCount")
    pElementCount = 0
    If pElementCount = 0 Then pEOA = True
End Sub
'////////////////////////////////////////////////////////////////////////////
Sub AddElementNext(ArgVariant As Variant)
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の追加エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementNext")
        GoTo ErrRtn
    End If
    Call AddElementBase(ArgVariant, pAbsolutePosition, "After")
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementNext")
End Sub
'////////////////////////////////////////////////////////////////////////////
Sub AddElementPrev(ArgVariant As Variant)
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の追加エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementPrev")
        GoTo ErrRtn
    End If
    Call AddElementBase(ArgVariant, pAbsolutePosition, "Before")
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementPrev")
End Sub
'////////////////////////////////////////////////////////////////////////////
Sub AddElementLast(ArgVariant As Variant)
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の追加エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementLast")
        GoTo ErrRtn
    End If
    Call AddElementBase(ArgVariant, pElementCount, "After")
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementLast")
End Sub
'////////////////////////////////////////////////////////////////////////////
Sub AddElementFirst(ArgVariant As Variant)
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の追加エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "AddElementFirst")
        GoTo ErrRtn
    End If
    Call AddElementBase(ArgVariant, 0, "Before")
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementFirst")
End Sub
'////////////////////////////////////////////////////////////////////////////
'Next
Private Sub AddElementBase( _
            ArgVariant As Variant, _
            AddElementPositionNumber As Long, _
            Before_After As String _
        )
On Error GoTo ErrRtn
    Dim AddDataVariant As Variant
    
    '========== Begin ==========
    'ArgVariantのチェック
    '変換して中継Variant型の変数に代入
    AddDataVariant = ConvertAndPutInReplayVariant(ArgVariant)
    'このVariantをArgvbTypeFlagInteger　任意の型に変換
    '↑の関数だけでは型判定を勝手にされてしまうため、意図しない型に入ることが予想される
    AddDataVariant = ConvertTypeOfVariantContents(AddDataVariant, pMainArrayTypeInteger)
    'プライベートのMain配列に追加する メイン関数
    Call AddElementDsignatePosition(AddDataVariant, AddElementPositionNumber, Before_After)
    
    Call UpdateElementCount
    'AbsolutePositoinチェック
        'EOAだった場合EOAではなくなる
    Call UpdatePositionStatus
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementNowPosition")
End Sub
'////////////////////////////////////////////////////////////////////////////
'値と要素を追加するときに追加する値がエラー値でないかチェック
'エラーなら空に
Private Function ConvertArgValueToVariantForAddElement( _
            ArgVariant As Variant _
        ) As Variant
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim RtnVariant As Variant
    Select Case VarType(ArgVariant)
            Case vbEmpty
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbEmpty]"
                Call SaveSystemError(-10, ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case vbNull
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbNull]"
                Call SaveSystemError(-vbNull, ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case vbInteger
                RtnVariant = CInt(ArgVariant)
            Case vbLong
                RtnVariant = CLng(ArgVariant)
            Case vbSingle
                RtnVariant = CSng(ArgVariant)
            Case vbDouble
                RtnVariant = CDbl(ArgVariant)
            Case vbCurrency
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbCurrency]"
                Call SaveSystemError(-vbCurrency, ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case vbDate
                RtnVariant = CDate(ArgVariant)
            Case vbString
                RtnVariant = CStr(ArgVariant)
            Case vbObject
                RtnVariant = (ArgVariant)
'                'エラー：セットできない
'                ErrorMessage = "ConvertArray Error"
'                ErrorMessage = ErrorMessage & "[vbObject]"
'                Call SaveSystemError(-vbObject, ErrorMessage, _
'                    "ConvertArgValueToVariantForAddElement" _
'                )
'                GoTo ErrRtn
            Case vbError
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbError]"
                Call SaveSystemError(-vbError, ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case vbBoolean
                RtnVariant = CBool(ArgVariant)
            Case vbVariant
                RtnVariant = CVar(ArgVariant)
            Case vbDataObject
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbDataObject]"
                Call SaveSystemError(-vbDataObject, ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case vbDecimal
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbDecimal]"
                Call SaveSystemError(-vbDecimal, ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case vbByte
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbByte]"
                Call SaveSystemError(-vbByte, ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case Is >= vbArray
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbArray]"
                Call SaveSystemError(-VarType(ArgVariant), ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
            Case Else
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + CaseElse]"
                Call SaveSystemError(-VarType(ArgVariant), ErrorMessage, _
                    "ConvertArgValueToVariantForAddElement" _
                )
                GoTo ErrRtn
    End Select
    ConvertArgValueToVariantForAddElement = RtnVariant
    '==========  End  ==========
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "ConvertArgValueToVariantForAddElement")
    ConvertArgValueToVariantForAddElement = RtnVariant
End Function
'////////////////////////////////////////////////////////////////////////////
Sub DebagPrintError()
On Error GoTo ErrRtn
    '========== Begin ==========
    If IsArrayExists(pErrorDescriptionArrayString) Then
        Dim i As Integer
        For i = 0 To UBound(pErrorDescriptionArrayString)
            Debug.Print "[" & pErrorNumberArrayInteger(i) & "] " & _
                pErrorDescriptionArrayString(i) & _
                 "(" & pErrorFunctionNameArrayString(i) & ")"
        Next i
    Else
    
    End If
    '==========  End  ==========
Exit Sub
ErrRtn:
    Debug.Print "[Cl_Array:DebagPrintError]"
    Debug.Print "[" & Err.Number & "]" & Err.Description
End Sub
'////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////
'MoveFirst レコードセツト内の先頭のレコード移動する
Sub MoveFirst()
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の移動エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "MoveFirst")
        GoTo ErrRtn
    End If
    Call Move(-pAbsolutePosition)
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "MoveFirst")
End Sub
'////////////////////////////////////////////////////////////////////////////
'MoveLast レコードセツト内の末尾のレコードに移動する
Sub MoveLast()
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の移動エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "MoveLast")
        GoTo ErrRtn
    End If
    Call Move((pElementCount - AbsolutePosition))
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "MoveLast")
End Sub
'////////////////////////////////////////////////////////////////////////////
'MoveNext レコードセツト内の次のレコードに移動する
Sub MoveNext()
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の移動エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "MoveNext")
        GoTo ErrRtn
    End If
    Call Move(1)
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "MoveNext")
End Sub
'////////////////////////////////////////////////////////////////////////////
'MovePrevious レコードセツト内の1つ前のレコードに移動する
Sub MovePrev()
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の移動エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "MovePrev")
        GoTo ErrRtn
    End If
    Call Move(-1)
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "MovePrev")
End Sub
'////////////////////////////////////////////////////////////////////////////
'Move 移動するレコード数を引数として指定できます｡
'引数が正の場合は末尾に向かって移動し､負の場合は先頭に向かって移動します｡
Sub Move(MoveNumber As Long)
On Error GoTo ErrRtn
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の移動エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "Move")
        GoTo ErrRtn
    End If
    
    pAbsolutePosition = pAbsolutePosition + MoveNumber
    If pAbsolutePosition > pElementCount Then
        pAbsolutePosition = pElementCount + 1
    End If
    Call UpdatePositionStatus
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "Move")
End Sub
'////////////////////////////////////////////////////////////////////////////
Sub DeleteElement(DeletePositionNumberLong As Long)
On Error GoTo ErrRtn
    Dim EditArrayVariant As Variant
    
    '========== Begin ==========
    If pMainArrayTypeInteger = 0 Then
        Dim ErrorMessage As String
        ErrorMessage = "配列未セット時の削除エラーです。"
        Call SaveSystemError(-10, ErrorMessage, "DeleteElement")
        GoTo ErrRtn
    End If
    
    Call DeleteBaseSingleElement(pMainArrayTypeInteger, DeletePositionNumberLong)
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "DeleteElement")
End Sub
'////////////////////////////////////////////////////////////////////////////
Sub DeleteBaseSingleElement( _
            ArgVbTypeInteger As Integer, _
            DeletePositionNumberLong As Long _
        )
On Error GoTo ErrRtn
    Dim EditArrayVariant As Variant
    
    '========== Begin ==========
    '変換して中継Variant型の変数に代入
    EditArrayVariant = GetMainArrayAfterConversionToVariant()
    '目的の配列の要素を一つ削除
    EditArrayVariant = DeleteSingleElementForArrayOfTypeVariant( _
        EditArrayVariant, ArgVbTypeInteger, DeletePositionNumberLong _
    )
    '変換した配列をもったVariantをPrivateのそれぞれの型の配列に入れる
    Call SetArrayMainMethod(EditArrayVariant)
    
    Call UpdateElementCount
    'AbsolutePositoinチェック
    Call UpdatePositionStatus
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "AddElementNowPosition")
End Sub
'////////////////////////////////////////////////////////////////////////////
'目的の配列の要素を一つ削除
Private Function DeleteSingleElementForArrayOfTypeVariant( _
            ArgDeleteArrayOfVariant As Variant, _
            ArgDeleteArrayTypeInteger As Integer, _
            ArgDeletePositionNumberLong As Long _
        ) As Variant
On Error GoTo ErrRtn
    
    Dim TempArrayInteger() As Integer
    Dim TempArrayLong() As Long
    Dim TempArraySingle() As Single
    Dim TempArrayDouble() As Double
'    Dim TempArrayCrrency() As Crrency
    Dim TempArrayDate() As Date
    Dim TempArrayString() As String
    Dim TempArrayBoolean() As Boolean
    Dim TempArrayObject() As Object
    '========== Begin ==========
    Dim ErrorMessage As String
    Dim Count As Long
    Dim ForVar As Variant
    For Each ForVar In ArgDeleteArrayOfVariant
        Select Case ArgDeleteArrayTypeInteger
            Case vbEmpty
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbEmpty]"
                Call SaveSystemError(-10, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbNull
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbNull]"
                Call SaveSystemError(-vbNull, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbInteger
                If Not ArgDeletePositionNumberLong = Count Then
                    If ArgDeletePositionNumberLong > Count Then
                        'DeletePositionより下　はそのまま
                        ReDim Preserve TempArrayInteger(Count)
                        TempArrayInteger(Count) = CInt(ForVar)
                    Else
                        'DeletePositionより上
                        ReDim Preserve TempArrayInteger(Count - 1)
                        TempArrayInteger(Count - 1) = CInt(ForVar)
                    End If
                Else
                    '同じ時は何もしない
                End If
            Case vbLong
                If Not ArgDeletePositionNumberLong = Count Then
                    If ArgDeletePositionNumberLong > Count Then
                        'DeletePositionより下　はそのまま
                        ReDim Preserve TempArrayLong(Count)
                        TempArrayLong(Count) = CLng(ForVar)
                    Else
                        'DeletePositionより上
                        ReDim Preserve TempArrayLong(Count - 1)
                        TempArrayLong(Count - 1) = CLng(ForVar)
                    End If
                Else
                    '同じ時は何もしない
                End If
            Case vbSingle
                If Not ArgDeletePositionNumberLong = Count Then
                    If ArgDeletePositionNumberLong > Count Then
                        'DeletePositionより下　はそのまま
                        ReDim Preserve TempArraySingle(Count)
                        TempArraySingle(Count) = CSng(ForVar)
                    Else
                        'DeletePositionより上
                        ReDim Preserve TempArraySingle(Count - 1)
                        TempArraySingle(Count - 1) = CSng(ForVar)
                    End If
                Else
                    '同じ時は何もしない
                End If
            Case vbDouble
                If Not ArgDeletePositionNumberLong = Count Then
                    If ArgDeletePositionNumberLong > Count Then
                        'DeletePositionより下　はそのまま
                        ReDim Preserve TempArrayDouble(Count)
                        TempArrayDouble(Count) = CDbl(ForVar)
                    Else
                        'DeletePositionより上
                        ReDim Preserve TempArrayDouble(Count - 1)
                        TempArrayDouble(Count - 1) = CDbl(ForVar)
                    End If
                Else
                    '同じ時は何もしない
                End If
            Case vbCurrency
'                ReDim Preserve TempArrayCrrency(Count)
'                TempArrayCrrency(Count) = CCur(ForVar)
'                If TypeFlagInteger = 0 Then TypeFlagInteger = vbCurrency
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbCurrency]"
                Call SaveSystemError(-vbCurrency, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbDate
                If Not ArgDeletePositionNumberLong = Count Then
                    If ArgDeletePositionNumberLong > Count Then
                        'DeletePositionより下　はそのまま
                        ReDim Preserve TempArrayDate(Count)
                        TempArrayDate(Count) = CDate(ForVar)
                    Else
                        'DeletePositionより上
                        ReDim Preserve TempArrayDate(Count - 1)
                        TempArrayDate(Count - 1) = CDate(ForVar)
                    End If
                Else
                    '同じ時は何もしない
                End If
            Case vbString
                If Not ArgDeletePositionNumberLong = Count Then
                    If ArgDeletePositionNumberLong > Count Then
                        'DeletePositionより下　はそのまま
                        ReDim Preserve TempArrayString(Count)
                        TempArrayString(Count) = CStr(ForVar)
                    Else
                        'DeletePositionより上
                        '1つずらす、ForEachで回しているのでループカウンタが扱いにくい
                        ReDim Preserve TempArrayString(Count - 1)
                        TempArrayString(Count - 1) = CStr(ForVar)
                    End If
                Else
                    '同じ時は何もしない
                End If
            Case vbObject
                If Not ArgDeletePositionNumberLong = Count Then
                    If ArgDeletePositionNumberLong > Count Then
                        'DeletePositionより下　はそのまま
                        ReDim Preserve TempArrayObject(Count)
                        Set TempArrayObject(Count) = (ForVar)
                    Else
                        'DeletePositionより上
                        '1つずらす、ForEachで回しているのでループカウンタが扱いにくい
                        ReDim Preserve TempArrayObject(Count - 1)
                        Set TempArrayObject(Count - 1) = (ForVar)
                    End If
                Else
                    '同じ時は何もしない
                End If
                
'                'エラー：セットできない
'                ErrorMessage = "ConvertArray Error"
'                ErrorMessage = ErrorMessage & "[vbArray + vbObject]"
'                Call SaveSystemError(-vbObject, ErrorMessage, "ConvertAndPutInReplayVariant")
'                GoTo ErrRtn
            Case vbError
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbError]"
                Call SaveSystemError(-vbError, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbBoolean
                If Not ArgDeletePositionNumberLong > Count Then
                    ReDim Preserve TempArrayBoolean(Count)
                    TempArrayBoolean(Count) = CBool(ForVar)
                Else
                    ReDim Preserve TempArrayBoolean(Count - 1)
                    TempArrayBoolean(Count - 1) = CBool(ForVar)
                End If
            Case vbVariant
'                ReDim Preserve TempArrayVariant(Count)
'                TempArrayVariant(Count) = CVar(ForVar)
'                If TypeFlagInteger = 0 Then TypeFlagInteger = vbVariant
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbVariant]"
                Call SaveSystemError(-vbVariant, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbDataObject
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbDataObject]"
                Call SaveSystemError(-vbDataObject, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbDecimal
                'エラー：セットできない
                ErrorMessage = "SetArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbDecimal]"
                Call SaveSystemError(-vbDecimal, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case vbByte
                'エラー：実装していません
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbByte]"
                Call SaveSystemError(-vbByte, ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case Is >= vbArray
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + vbArray]"
                Call SaveSystemError(-VarType(ArgDeleteArrayOfVariant), ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
            Case Else
                'エラー：セットできない
                ErrorMessage = "ConvertArray Error"
                ErrorMessage = ErrorMessage & "[vbArray + CaseElse]"
                Call SaveSystemError(-VarType(ArgDeleteArrayOfVariant), ErrorMessage, "ConvertAndPutInReplayVariant")
                GoTo ErrRtn
                'vbArray + NUMBER
        End Select
        Count = Count + 1
    Next
    '一度変換したものを ReturnVariant へ
    Dim ReturnVariant As Variant
    Select Case ArgDeleteArrayTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                ReturnVariant = TempArrayInteger
            Case vbLong
                ReturnVariant = TempArrayLong
            Case vbSingle
                ReturnVariant = TempArraySingle
            Case vbDouble
                ReturnVariant = TempArrayDouble
            Case vbCurrency
'                ReturnVariant = TempArrayCrrency
            Case vbDate
                ReturnVariant = TempArrayDate
            Case vbString
                ReturnVariant = TempArrayString
            Case vbObject
                ReturnVariant = TempArrayObject
            Case vbError
            Case vbBoolean
                ReturnVariant = TempArrayBoolean
            Case vbVariant
'                ReturnVariant = TempArrayVariant
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    DeleteSingleElementForArrayOfTypeVariant = ReturnVariant
    '==========  End  ==========
Exit Function
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "DeleteSingleElementForArrayOfTypeVariant")
    DeleteSingleElementForArrayOfTypeVariant = ReturnVariant
End Function
'////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArrayInteger() As Integer
On Error GoTo ErrRtn
    '========== Begin ==========
    GetSingleElementOfArrayInteger = pMainArrayInteger(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  ==========
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArrayInteger")
    GetSingleElementOfArrayInteger = "0"
End Function
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArrayLong() As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    GetSingleElementOfArrayLong = pMainArrayLong(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  ==========
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArrayLong")
    GetSingleElementOfArrayLong = "0"
End Function
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArraySingle() As Single
On Error GoTo ErrRtn
    '========== Begin ==========
    GetSingleElementOfArraySingle = pMainArraySingle(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  ==========
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArraySingle")
    GetSingleElementOfArraySingle = "0"
End Function
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArrayDouble() As Double
On Error GoTo ErrRtn
    '========== Begin ==========
    GetSingleElementOfArrayDouble = pMainArrayDouble(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  ==========
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArrayDouble")
    GetSingleElementOfArrayDouble = "0"
End Function
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArrayDate() As Date
On Error GoTo ErrRtn
    '========== Begin ==========
    GetSingleElementOfArrayDate = pMainArrayDate(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  ==========
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArrayDate")
    GetSingleElementOfArrayDate = "0"
End Function
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArrayBoolean() As Boolean
On Error GoTo ErrRtn
    '========== Begin ==========
    GetSingleElementOfArrayBoolean = pMainArrayBoolean(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  ==========
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArrayBoolean")
    GetSingleElementOfArrayBoolean = "0"
End Function
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArrayString() As String
On Error GoTo ErrRtn
    '========== Begin ==========
    GetSingleElementOfArrayString = pMainArrayString(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  ==========
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArrayString")
    GetSingleElementOfArrayString = "__ERROR__"
End Function
'////////////////////////////////////////////////////////////////////////////
Function GetSingleElementOfArrayObject() As Object
On Error GoTo ErrRtn
    '========== Begin ==========
    Set GetSingleElementOfArrayObject = pMainArrayObject(pAbsolutePosition)
    Call CheckValueOrTheSameWhenGetSingleElement
    '==========  End  === ======
Exit Function
ErrRtn:
Dim ErrorMessage As String
ErrorMessage = "[Position=" & pAbsolutePosition & "]"
Call SaveSystemError(Err.Number, Err.Description & ErrorMessage, "GetSingleElementOfArrayObject")
   Set GetSingleElementOfArrayObject = Nothing
End Function
'////////////////////////////////////////////////////////////////////////////
Private Sub CheckValueOrTheSameWhenGetSingleElement()
On Error GoTo ErrRtn
    Dim CheckValueVariant As Variant
    '========== Begin ==========
    Select Case pMainArrayTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                CheckValueVariant = pMainArrayInteger(pAbsolutePosition)
            Case vbLong
                CheckValueVariant = pMainArrayLong(pAbsolutePosition)
            Case vbSingle
                CheckValueVariant = pMainArraySingle(pAbsolutePosition)
            Case vbDouble
                CheckValueVariant = pMainArrayDouble(pAbsolutePosition)
            Case vbCurrency
                CheckValueVariant = pMainArrayCrrency(pAbsolutePosition)
            Case vbDate
                CheckValueVariant = pMainArrayDate(pAbsolutePosition)
            Case vbString
                CheckValueVariant = pMainArrayString(pAbsolutePosition)
            Case vbObject
                Set CheckValueVariant = pMainArrayObject(pAbsolutePosition)
            Case vbError
            Case vbBoolean
                CheckValueVariant = pMainArrayBoolean(pAbsolutePosition)
            Case vbVariant
'                CheckValueVariant = pMainArrayVariant(pAbsolutePosition)
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    If pMainArrayTypeInteger = vbObject Then
        CheckValueVariant = pAbsolutePosition
    End If
    
    If pValueWhenGetSameValueVariant = CheckValueVariant Then
        pCountIntegerWhenGetSameValue = pCountIntegerWhenGetSameValue + 1
        If pCountIntegerWhenGetSameValue >= pMaxCountIntegerWhenGetSameValue Then
            '-1は上限なし
            If Not pMaxCountIntegerWhenGetSameValue = -1 Then
                AbsolutePosition = pElementCount + 1
                Dim ErrorMessage As String
                'エラー：セットできない
                ErrorMessage = "同じ値が" & pMaxCountIntegerWhenGetSameValue & "回参照されたので、EOAに移動します。"
    '            ErrorMessage = ErrorMessage & "[vbArray + vbDataObject]"
                Call SaveSystemError(-10, ErrorMessage, "CheckValueOrTheSameWhenGetSingleElement")
                GoTo ErrRtn
            End If
        End If
    Else
        pValueWhenGetSameValueVariant = CheckValueVariant
    End If
    Set CheckValueVariant = Nothing
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "CheckValueOrTheSameWhenGetSingleElement")
    Set CheckValueVariant = Nothing
End Sub
'////////////////////////////////////////////////////////////////////////////
Private Sub CheckArrayStatus()
On Error GoTo ErrRtn
    Dim CheckValueVariant As Variant
    '========== Begin ==========
    Dim ExistsFlagBoolean As Boolean
    Select Case pMainArrayTypeInteger
            Case vbEmpty
            Case vbNull
            Case vbInteger
                ExistsFlagBoolean = IsArrayExists(pMainArrayInteger)
            Case vbLong
                ExistsFlagBoolean = IsArrayExists(pMainArrayLong)
            Case vbSingle
                ExistsFlagBoolean = IsArrayExists(pMainArraySingle)
            Case vbDouble
                ExistsFlagBoolean = IsArrayExists(pMainArrayDouble)
            Case vbCurrency
                ExistsFlagBoolean = IsArrayExists(pMainArrayCrrency)
            Case vbDate
                ExistsFlagBoolean = IsArrayExists(pMainArrayDate)
            Case vbString
                ExistsFlagBoolean = IsArrayExists(pMainArrayString)
            Case vbObject
                ExistsFlagBoolean = IsArrayExists(pMainArrayObject)
            Case vbError
            Case vbBoolean
                ExistsFlagBoolean = IsArrayExists(pMainArrayBoolean)
            Case vbVariant
'                CheckValueVariant = pMainArrayVariant(pAbsolutePosition)
            Case vbDataObject
            Case vbDecimal
            Case vbByte
            Case Is >= vbArray
            Case Else
    End Select
    
    If ExistsFlagBoolean = False Then
        pEOA = True
        pBOA = True
        pAbsolutePosition = 0
        pElementCount = 0
    End If
    '==========  End  ==========
Exit Sub
ErrRtn:
Call SaveSystemError(Err.Number, Err.Description, "CheckValueOrTheSameWhenGetSingleElement")
End Sub
'////////////////////////////////////////////////////////////////////////////



    'SetArray  Error
    
    'Same Value


    'Flagすべての型の準備
    'オブジェクト型は使えるか？
    '========
    
    



    'レコードが削除されたとき注意
'AbsolutePosition
    '追加されたとき
    '削除されたとき
    'クリアされたとき
    'Setされたとき
    'Moveされたとき

'MaxCount が変わるとき
    'ADDされたとき
    'Delされたとき
    'クリアされたとき
    'Setされたとき
        
'BOA EOA が変わるとき  Absolute Positionが変更されたときに一緒に処理する
    '追加されたとき
        
    '削除されたとき
    'クリアされたとき
    'Setされたとき
    'Moveされたとき
    
'Last から MoveNext -> EOA
'First から MovePrev -> BOA

    
'////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////
'find array string
    '方向　一致部分

'Write は要らない
    

'Chk Array ->[ [Type] \n 001:aaa   002:bbb ... ]
    'Debug.print   OutPutInCell(Sheet,Address,SingleCell_MultiCells)

'Array
    'String
        'Length 要素数 size CountElements
        'CountValue '文字列配列の中に空文字がいくつあるか
        'AddElements
        '特定記号などで文字列をsplitして、argNumber番目の文字列を返す
            '↑これから整数などほかの方にしたい場合は　Cnvで変換




    'Extention


