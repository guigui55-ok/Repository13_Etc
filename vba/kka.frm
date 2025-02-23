VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7590
   OleObjectBlob   =   "kka.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    '検索実行
    Call FindMain
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm1
End Sub

Private Sub CommandButton3_Click()
    Call GetClipBord
End Sub

Sub GetClipBord()
    'DataObjectオブジェクトはMSFormsのメンバです。使用するには、
    'Microsoft Forms 2.0 Object Libraryを参照設定
    
    Dim CB As DataObject
    Set CB = New DataObject
'    buf = "tanaka"
    With CB
'        .SetText buf        ''変数のデータをDataObjectに格納する
'        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
'        buf2 = .GetText     ''DataObjectのデータを変数に取得する
        
    End With
'    MsgBox buf2
TextBox1.Value = CB.GetText
Set CB = Nothing

End Sub



Private Sub CommandButton4_Click()
'    If a Then
    Application.WindowState = xlMinimized
End Sub

Private Sub TextBox1_Change()

End Sub

Sub FindMain()
    Dim FlagLong As Long
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim cCaptionAddressList As Cl_Array
    Set cCaptionAddressList = New Cl_Array
    Dim cSheetNameList As Cl_Array
    Set cSheetNameList = New Cl_Array
    Dim cFieldNameList As Cl_Array
    Set cFieldNameList = New Cl_Array
    Dim cFieldCaptionAddressList As Cl_Array
    Set cFieldCaptionAddressList = New Cl_Array
    Dim ResultArrayString() As String
    Dim TextBoxValue As String
    
    TextBoxValue = TextBox1.Value
    
    'グローバルからプライベートへ
    Call cCaptionAddressList.SetArray(gCaptionAddressList, vbString)
    Call cSheetNameList.SetArray(gSheetList, vbString)
    Call cFieldNameList.SetArray(gFielsNameList, vbString)
    
    'ループ対策の同値参照 上限回数の解除設定
    cSheetNameList.MaxCountIntegerWhenGetSameValue = -1
    
    'テキストボックスが空なら終了する
    If TextBoxValue = "" Then
        ShowMsgBox "TextBox is Blank"
        GoTo ErrRtn
    End If
    Logout ("TextBoxValue = " + TextBoxValue)

    'テーブルの個数分ループ
    Do While (Not cCaptionAddressList.EOA)
        '
        Dim ResultSingleTableArrayString() As String
        Dim FieldCaptionAddress As String
        Dim FindCompareFieldCaptionAddress As String  'Compare = 比較する
        Dim CompareBeginAddress As String
        Dim CompareTimes As Long
'        Dim cFieldCaptionAddressList As New Cl_Array
        
'        DP (cSheetNameList.GetDirectMainArrayString): Stop
'        Debug.Print cSheetNameList.GetSingleElementOfArrayString
        
        'キャプションアドレス取得
        'フィールド名リストのアドレスを取得設定
        FieldCaptionAddress = GetFieldLabelAddress( _
            cSheetNameList.GetSingleElementOfArrayString, _
            cCaptionAddressList.GetSingleElementOfArrayString _
        )
        If Not IsAddress(FieldCaptionAddress) Then
            Debug.Print "Error=-210"
            GoTo ErrRtn
        End If
        '値保存用フィールド名アドレスリストを作成しておく
        cFieldCaptionAddressList.SetDirectMainArrayString = _
            SetFieldCaptionAddressList( _
                cSheetNameList.GetSingleElementOfArrayString, _
                FieldCaptionAddress, _
                cFieldNameList.GetDirectMainArrayString _
        )
'        DP (cFieldCaptionAddressList.GetDirectMainArrayString): Stop
        
        
        '比較要素のCol取得
        FindCompareFieldCaptionAddress = GetColumnCompareFieldName( _
            cSheetNameList.GetSingleElementOfArrayString, _
            FieldCaptionAddress, _
            "ファイル名" _
        )
        If Not IsAddress(FindCompareFieldCaptionAddress) Then
            Debug.Print "Error=-220"
            Debug.Print cSheetNameList.GetSingleElementOfArrayString & " / " & cCaptionAddressList.GetSingleElementOfArrayString
            GoTo ErrRtn
        End If
        'Set Begin Address
        CompareBeginAddress = SetBeginAddress( _
            cSheetNameList.GetSingleElementOfArrayString, _
            cCaptionAddressList.GetSingleElementOfArrayString, _
            FindCompareFieldCaptionAddress _
        )
        If Not IsAddress(CompareBeginAddress) Then
            Debug.Print "Error=-230"
            GoTo ErrRtn
        End If
        'EndRow　　'回数
        CompareTimes = SetCompareTimes( _
            cSheetNameList.GetSingleElementOfArrayString, _
            CompareBeginAddress _
        )
        If CompareTimes < 1 Then
            Debug.Print "Error=-240"
            GoTo ErrRtn
        End If
        '行数分ループ
        Dim i As Long
        For i = 0 To CompareTimes
            '[ファイル名]の文字列を比較
            FlagLong = CompareValue( _
                TextBoxValue, _
                cSheetNameList.GetSingleElementOfArrayString, _
                CompareBeginAddress, _
                i _
            )
            'あれば必要フィールド値をまとめて保存
            If FlagLong > 0 Then
                FlagLong = SaveValues( _
                    cFieldCaptionAddressList, _
                    cSheetNameList.GetSingleElementOfArrayString, _
                    CompareBeginAddress, _
                    i, _
                    ResultArrayString _
                )
            Else
                'ない
            End If
        Next
        '次へ
        cCaptionAddressList.MoveNext
        cSheetNameList.MoveNext
    Loop
    Dim cResultArray As New Cl_Array
    Call cResultArray.SetArray(ResultArrayString, vbString)
'    DP (ResultArrayString): Stop
    If IsArrayExists(ResultArrayString) Then
        gResultData = ResultArrayString
'        Do While (Not cResultArray.EOA)
'            TextBox2.Text = TextBox2.Text & cResultArray.GetSingleElementOfArrayString & vbNewLine
'            cResultArray.MoveNext
'        Loop
        UserForm2.Show vbModeless
'        TextBoxValue = TextBox1.Value
'        TextBox2.はEnterKeyBehavior = True
        'Enterキーで改行するに はEnterKeyBehavior をTrueに設定します。
    Else
        MsgBox "検索条件に一致する項目はありません。"
    End If
    '==========  End  ==========
    Set cCaptionAddressList = Nothing
    Set cSheetNameList = Nothing
    Set cFieldNameList = Nothing
    Set cFieldCaptionAddressList = Nothing
    Erase ResultArrayString
FlagLong = 1
Exit Sub
ErrRtn:
    Set cCaptionAddressList = Nothing
    Set cSheetNameList = Nothing
    Set cFieldNameList = Nothing
    Set cFieldCaptionAddressList = Nothing
    Erase ResultArrayString
    DPErr
End Sub

'あれば必要フィールド値をまとめて保存
Function SaveValues( _
        cFieldAddressList As Cl_Array, _
        ArgSheetName As String, _
        BeginAddress As String, _
        LoopCount As Long, _
        ByRef SaveItemArrayString() As String _
    ) As Long
    '===========================
    Dim FlagLong As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim cTempArray As New Cl_Array
    Dim GetInfoString As String
    Dim GetRow As Long
    Dim GetCol As Long
    Dim TempString As String
    
    cFieldAddressList.MoveFirst
    Do While (Not cFieldAddressList.EOA)
        '最初以外は区切りを
        If Not cFieldAddressList.AbsolutePosition = 0 Then
            GetInfoString = GetInfoString & " / "
        End If
        'フィールドラベルがないときがある
        If Not cFieldAddressList.GetSingleElementOfArrayString = "" Then
            GetRow = Range(BeginAddress).Offset(LoopCount, 0).Row
            GetCol = Range(cFieldAddressList.GetSingleElementOfArrayString).Column
            
            '評価の桁数をそろえる
            If cFieldAddressList.AbsolutePosition = 0 Then
                TempString = Sheets(ArgSheetName).Cells(GetRow, GetCol).Value
                Dim SpaceCount As Integer
                SpaceCount = 10 - LenB(TempString)
                If SpaceCount > 0 Then
                    Dim i As Integer
                     For i = 0 To SpaceCount
                        TempString = TempString & " "
                     Next i
                End If
                GetInfoString = GetInfoString & _
                    TempString
            Else
                'Position > 0
                GetInfoString = GetInfoString & _
                    Sheets(ArgSheetName).Cells(GetRow, GetCol).Value
            End If
            
        End If
        
        cFieldAddressList.MoveNext
    Loop
    'すべてまとめたら　配列に追記
    SaveItemArrayString = cTempArray.GetValueDirect _
        .StringAddValueLastElements_RtnString( _
            SaveItemArrayString, GetInfoString _
    )
'    DP (SaveItemArrayString): Stop
    '==========  End  ==========
    Set cTempArray = Nothing
    SaveValues = FlagLong
Exit Function
    Set cTempArray = Nothing
ErrRtn:
    
    SaveValues = -1
End Function



'[ファイル名]の文字列を比較
Function CompareValue( _
        TextBoxValue As String, _
        ArgSheetName As String, _
        BeginAddress As String, _
        LoopCount As Long _
    ) As Long
    '===========================
    Dim FlagLong As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim GetValue As String
    GetValue = Sheets(ArgSheetName).Range(BeginAddress).Offset(LoopCount, 0).Value
    
    'CompareValue の中に TextBoxValue が含まれていればOK
    If InStr(1, GetValue, TextBoxValue, vbBinaryCompare) > 0 Then
        FlagLong = 1
    Else
        FlagLong = -2
    End If
    '==========  End  ==========
    CompareValue = FlagLong
Exit Function
ErrRtn:
    DPErr
    CompareValue = -1
End Function


'回数
Function SetCompareTimes( _
        ArgSheetName As String, _
        BeginAddress As String _
    ) As Long
    '===========================
    Dim CompareTimes As Long
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim EndRow As Long
    Dim cTempCell As New Cl_Cell
    
    Call cTempCell.SetSheetAddress(ArgSheetName, BeginAddress)
    EndRow = cTempCell.GetValue.EndxlDownRowLongLastOfContinuousCells
    
    CompareTimes = EndRow - Range(BeginAddress).Row
    '==========  End  ==========
    Set cTempCell = Nothing
    SetCompareTimes = CompareTimes
Exit Function
ErrRtn:
    Set cTempCell = Nothing
    SetCompareTimes = -1
End Function




'検索比較の最初のアドレス
Function SetBeginAddress( _
        ArgSheetName As String, _
        ArgTableCaptionAddress As String, _
        ArgFindCompareFieldAddress As String _
    ) As String
    '===========================
    Dim RtnAddress As String
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim BeginCol As Long
    Dim BeginRow As Long
    
    BeginCol = Range(ArgFindCompareFieldAddress).Column
    BeginRow = Range(ArgTableCaptionAddress).Offset(2, 0).Row
    
    RtnAddress = Cells(BeginRow, BeginCol).Address
    '==========  End  ==========
    SetBeginAddress = RtnAddress
Exit Function
ErrRtn:
    SetBeginAddress = "__ERROR"
End Function


'比較要素のCol取得
Function GetColumnCompareFieldName( _
        ArgSheetName As String, _
        ArgFieldCaptionAddress As String, _
        ArgFindMainField As String _
    ) As String
    '===========================
    Dim RtnAddress As String
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim cTempCell As New Cl_Cell
    
    Call cTempCell.SetSheetAddress(ArgSheetName, ArgFieldCaptionAddress)
    RtnAddress = cTempCell.GetValue.FindString_RtnAddress( _
        ArgSheetName, _
        cTempCell.GetValue.CnvRangeAddressToSigleAddresOfStart(ArgFieldCaptionAddress), _
        ArgFieldCaptionAddress, _
        ArgFindMainField, _
        xlWhole, 0, xlNext, False, False _
    )
    
        
    '==========  End  ==========
    Set cTempCell = Nothing
    GetColumnCompareFieldName = RtnAddress
Exit Function
ErrRtn:
    Set cTempCell = Nothing
    GetColumnCompareFieldName = "__ERROR"
End Function

'値保存用フィールド名アドレスリストを作成しておく
Function SetFieldCaptionAddressList( _
        ArgSheetName As String, _
        ArgFieldCaptionRangeAddress As String, _
        ArgFieldNameList() As String _
    ) As String()
    '===========================
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim cTempCell As New Cl_Cell
    Dim cTempArray As New Cl_Array
    Dim ReturnAddressList() As String
    
    Call cTempCell.SetSheetAddress(ArgSheetName, ArgFieldCaptionRangeAddress)
    Dim i As Integer
    Dim FieldCaptionSingleAddress As String
    For i = 0 To UBound(ArgFieldNameList)
        FieldCaptionSingleAddress = cTempCell.GetValue.FindString_RtnAddress( _
            ArgSheetName, _
            cTempCell.GetValue.CnvRangeAddressToSigleAddresOfStart(ArgFieldCaptionRangeAddress), _
            ArgFieldCaptionRangeAddress, _
            ArgFieldNameList(i), _
            xlWhole, 0, xlNext, False, False _
        )
        '追加
        ReturnAddressList = cTempArray.GetValueDirect _
            .StringAddValueLastElements_RtnString( _
                ReturnAddressList, FieldCaptionSingleAddress _
        )
    Next i
    '==========  End  ==========
    Set cTempArray = Nothing
    Set cTempCell = Nothing
    SetFieldCaptionAddressList = ReturnAddressList
Exit Function
ErrRtn:
    Set cTempArray = Nothing
    Set cTempCell = Nothing
    SetFieldCaptionAddressList = ReturnAddressList
End Function


'フィールド名のキャプションアドレスを取得設定
Function GetFieldLabelAddress( _
        ArgSheetName As String, _
        ArgTableCaptionAddress As String _
    ) As String
    '===========================
    Dim FieldLabelAddress As String
On Error GoTo ErrRtn
    '========== Begin ==========
    Dim FieldLabelBeginAddress As String
    Dim FieldLabelEndAddress As String
    Dim cTempCell As New Cl_Cell
    '<<リスト>>から1つ下がフィールドBegin
    FieldLabelBeginAddress = Range(ArgTableCaptionAddress).Offset(1, 0).Address
    Call cTempCell.SetSheetAddress(ArgSheetName, FieldLabelBeginAddress)
    'その右が　FieldEnd
    FieldLabelEndAddress = cTempCell.GetValue _
        .EndxlRightLastOfContinuousCellsAddress
        
    FieldLabelAddress = FieldLabelBeginAddress & ":" & FieldLabelEndAddress
    
    If Not IsAddress(FieldLabelAddress) Then
        GoTo ErrRtn
    End If
    '==========  End  ==========
    Set cTempCell = Nothing
    GetFieldLabelAddress = FieldLabelAddress
Exit Function
ErrRtn:
    Set cTempCell = Nothing
    GetFieldLabelAddress = "__ERROR"
End Function
