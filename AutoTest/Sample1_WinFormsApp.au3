; =============================
; WinForms アプリ自動操作テンプレート
; =============================

; アプリを起動
Run("C:\Users\OK\source\repos\Repository10_VBnet\UiTestTryApp\UiTestTryApp\bin\Debug\UiTestTryApp.exe")

; ウィンドウのタイトルが現れるまで待機
WinWaitActive("UiTestTryApp")

; ボタンをクリック（例: OKボタン）
; ControlClick("window_title", "text", "control_id" [, button [, clicks [, x [, y]]]])
ControlClick("UiTestTryApp", "", "[CLASS:WindowsForms10.BUTTON.app.0.141b42a_r9_ad1; INSTANCE:1]")

; テキストボックスに入力
; ControlSetText("ウィンドウ名", "", "テキストボックスID", "入力内容")
ControlSetText("UiTestTryApp", "", "[CLASS:WindowsForms10.EDIT.app.0.141b42a_r9_ad1; INSTANCE:1]", "テスト入力")

; Enterキーを送信
Send("{ENTER}")
