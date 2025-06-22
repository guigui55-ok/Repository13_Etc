; =============================
; WinForms アプリ 自動操作 + ログ出力
; =============================

; ログファイル名の生成：Log_yyyyMMdd_HHmmss.log
Local $now = @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC
Local $logFileName = "Log_" & $now & ".log"
Local $logFilePath = @ScriptDir & "\" & $logFileName

; ログファイルを開く（2 = 書き込み、上書き）
Local $logFile = FileOpen($logFilePath, 2)
If $logFile = -1 Then
    MsgBox(16, "エラー", "ログファイルを開けません: " & $logFilePath)
    Exit
EndIf

; ログ書き込み関数
Func LogWrite($msg)
    FileWriteLine($logFile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "] " & $msg)
EndFunc

; アプリ起動
LogWrite("アプリ起動開始")
Run("C:\Users\OK\source\repos\Repository10_VBnet\UiTestTryApp\UiTestTryApp\bin\Debug\UiTestTryApp.exe")

; ウィンドウが表示されるまで待機
If WinWaitActive("UiTestTryApp", "", 5) Then
    LogWrite("ウィンドウを検出しました")
Else
    LogWrite("ウィンドウが見つかりません（5秒タイムアウト）")
    FileClose($logFile)
    Exit
EndIf

; ボタンクリック
LogWrite("ボタンをクリックします")
Local $ret1 = ControlClick("UiTestTryApp", "", "[TEXT:Execute]")
LogWrite("ControlClick結果: " & $ret1)

; テキスト入力
LogWrite("テキストを入力します")
Local $ret2 = ControlSetText("UiTestTryApp", "", "[NAME:TextBox1]", "テスト入力")
LogWrite("ControlSetText結果: " & $ret2)

; Enterキー送信
LogWrite("Enterキーを送信")
Send("{ENTER}")

; ログファイルを閉じる
LogWrite("処理完了")
FileClose($logFile)
