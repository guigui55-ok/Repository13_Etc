Attribute VB_Name = "Module8"
'Option Compare Database 'error
Option Explicit
  
Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
       (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
  
Private Const WM_VSCROLL    As Long = &H115&  ' 垂直スクロール
Private Const SB_LINEUP     As Long = &H0&    ' 上方向
Private Const SB_LINEDOWN   As Long = &H1&    ' 下方向
  
' [ACC2002] スクロールすると先頭レコードが表示されない
' http://support.microsoft.com/kb/418706/ja
' 上記現象の対策として、マウスホイール使用時イベントを使うサンプル
Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
  
    ' 下にスクロールした場合は関係ないので、終了
    If Count > 0 Then Exit Sub
  
    Dim MaxVisibleRecordCount   As Integer  ' 1 画面に表示可能な最大レコード数
    Dim iDetailsH               As Integer  ' 詳細セクション全体の高さ
    Dim iDetailH                As Integer  ' 詳細セクション一つ分の高さ
    Dim iHeaderH                As Integer  ' フォームヘッダーの高さ
    Dim iFooterH                As Integer  ' フォームフッターの高さ
    Dim iCurrAboveRows          As Integer  ' カレント行の上に表示されている行数
  
On Error Resume Next    ' フォームヘッダー/フッターが存在しない場合のエラーを無視
    iHeaderH = Me.Section(acHeader).Height
    iFooterH = Me.Section(acFooter).Height
On Error GoTo erh
    iDetailH = Me.Section(acDetail).Height
    iDetailsH = Me.InsideHeight - iHeaderH - iFooterH
    MaxVisibleRecordCount = iDetailsH \ iDetailH
    ' 追加が許可されている場合は、新規レコード 1 行分を補正
    If Me.AllowAdditions Then MaxVisibleRecordCount = MaxVisibleRecordCount - 1
  
    ' 全レコードを 1 画面に表示可能な場合
    ' -- 実際には、MaxVisibleRecordCount = Me.Recordset.RecordCount の
    '    場合、ホイールスクロールで 1 行目を表示できることもありますが、
    '    境界判定が困難なため、同値の場合を含めて対策します。
    If MaxVisibleRecordCount >= Me.Recordset.RecordCount Then
        iCurrAboveRows = (Me.CurrentSectionTop - iHeaderH) / iDetailH
  
        ' 上に隠れている非表示レコードが 1 行以下の場合(バグに該当)
        If (Me.CurrentRecord - iCurrAboveRows - 1) <= 1 Then
            SendMessage Me.hWnd, WM_VSCROLL, SB_LINEUP, 0&  ' 強制スクロール!
        End If
    End If
    Exit Sub
  
erh:
    MsgBox Err.Description, vbCritical, "実行時エラー" & Err.Number
End Sub

