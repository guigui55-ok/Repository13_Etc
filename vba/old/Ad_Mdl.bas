Attribute VB_Name = "Ad_Mdl"


'########################################################################################
'Globals
'########################################################################################
Dim gConst As ConstFindSystem

'グローバル変数
'テーブルのキャプションのアドレスリスト
Public gCaptionAddressList() As String
'シート名リスト
Public gSheetList() As String
'フィールド名リスト
Public gFielsNameList() As String
'検索結果リスト
Public gResultData() As String
'デバッグモード
'ONの時は、ログなどを出力する
Dim gDebugMode As Integer
'ログ出力用シート名の設定
Dim gLogoutSheetName As Integer
'ログ出力用セルアドレスの設定
Dim gLogoutBeginAddress As String
'ログ出力用アドレス
'現在出力している場所
Dim gLogoutAddress As String
'ログ出力用ファイルパス
Dim gLogoutPath As String
'ログ出力用インデント
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
'デバッグモードONのときセルへ出力する
Sub Logout(Value As Variant)
On Error GoTo ErrRtn
    'OFFなら終了する
    'DEBUG_MODE_ON=1
    If gDebugMode < 1 Then
        Exit Sub
    End If
    Dim buf As String
    buf = CnvVarToStr(Value)
    buf = Str(Now()) + " " + buf
    'イミディエイトへ出力する
    'LOG_TO_IMMIDIATE=2
    If gDebugMode And 2 Then
        ShowDebugPrint buf
    End If
    'セルに出力する
    'LOG_TO_CELL=4
    If gDebugMode And 4 Then
        'gLogoutPath
        Sheets(gLogoutSheetName).Range(gLogoutAddress).Value = buf
    End If
    'ファイルに出力する
    'LOG_TO_FILE=8
    If gDebugMode And 8 Then
        'gLogoutPath
        Flag = WriteFile(gLogoutPath, buf)
    End If
Exit Sub
ErrRtn:
    Debug.Print ("Module:Ad_Mdl,Function:Logout,Err=" + Err.Number + ":" + Err.Description)
End Sub


'デバッグモードONのときイミディエイトへ出力する
Sub ShowDebugPrint(Value As String)
    If gDebugMode >= DEBUG_ON Then
        Debug.Print (Value)
    End If
End Sub

'メッセージボックスをを表示する、デバッグモードONのときイミディエイトへも出力する
Sub ShowDebugMsgBox(Value As String)
    MsgBox Value
    If gDebugMode >= DEBUG_ON Then
        Debug.Print (Value)
    End If
End Sub


'########################################################################################
'テーブル上のデータを検索するシステム メイン実行関数
'########################################################################################
'########################################################################################
'複数の検索テーブルから文字列を検索する
'実行メソッド
Sub FindStringOfMultiTable(Optional debugMode As Integer = 0)
    Dim cFindString As AdCl_FindStringMain
    Set cFindString = New AdCl_FindStringMain
    
    '実行開始時の情報をログへ出力する
    Debug.Print ("Module:Ad_Mdl , Function:FindStringOfMultiTable , DebugMode:" + Str(debugMode))
    'デバッグモード変数をグローバルとクラスメンバへ格納する
    gDebugMode = debugMode
    cFindString.debugMode = debugMode
    '機能を実行する
    Call cFindString.Main
    
    Set cFindString = Nothing
End Sub

