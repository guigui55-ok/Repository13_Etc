; =============================
; WinForms アプリ自動操作テンプレート
; =============================

; アプリを起動
Run("C:\Users\OK\source\repos\Repository10_VBnet\UiTestTryApp\UiTestTryApp\bin\Debug\UiTestTryApp.exe")

; ウィンドウのタイトルが現れるまで待機
WinWaitActive("UiTestTryApp")

If WinWaitActive("UiTestTryApp", "", 5) Then
    ConsoleWrite("ウィンドウが表示されました。" & @CRLF)
Else
    ConsoleWrite("ウィンドウが見つかりません。" & @CRLF)
EndIf

ConsoleWrite("ボタンをクリックします..." & @CRLF)
Local $ret1 = ControlClick("UiTestTryApp", "", "[CLASS:WindowsForms10.BUTTON.app.0.141b42a_r9_ad1; INSTANCE:1]")
ConsoleWrite("ボタンクリック結果: " & $ret1 & @CRLF)

ConsoleWrite("テキストを入力します..." & @CRLF)
Local $ret2 = ControlSetText("UiTestTryApp", "", "[CLASS:WindowsForms10.EDIT.app.0.141b42a_r9_ad1; INSTANCE:1]", "テスト入力")
ConsoleWrite("テキスト入力結果: " & $ret2 & @CRLF)

; Enterキーを送信
Send("{ENTER}")

Local $logFile = FileOpen(@ScriptDir & "\test_log.txt", 2) ; 2 = 書き込みモード
FileWriteLine($logFile, "Window check start...")

If WinWaitActive("UiTestTryApp", "", 5) Then
    FileWriteLine($logFile, "Window is active")
Else
    FileWriteLine($logFile, "Window not found")
EndIf

FileWriteLine($logFile, "Trying button click...")
Local $ret1 = ControlClick("UiTestTryApp", "", "[CLASS:WindowsForms10.BUTTON.app.0.141b42a_r9_ad1; INSTANCE:1]")
FileWriteLine($logFile, "Click result: " & $ret1)

FileWriteLine($logFile, "Trying text input...")
Local $ret2 = ControlSetText("UiTestTryApp", "", "[CLASS:WindowsForms10.EDIT.app.0.141b42a_r9_ad1; INSTANCE:1]", "テスト入力")
FileWriteLine($logFile, "Text result: " & $ret2)

FileClose($logFile)