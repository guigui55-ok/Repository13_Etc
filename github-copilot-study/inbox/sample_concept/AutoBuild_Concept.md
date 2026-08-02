# AutoBuildのコンセプト

# 目的
VisualStudioで複数のソリューションをReleaseビルドして、
他のフォルダ（共有フォルダ）にコピーするのを自動化する。

# 仕様
- ソリューションファイルを指定して、ビルドする。
- ビルド後、指定したフォルダにコピーする。
- ビルドのログを保存する。
- ビルドの成功/失敗を通知する。
- スケジュールは設定しない。
- ビルド前のチェックなども行わない。
- ビルドの失敗は、ログを確認して手動で対応する。
- ビルドの成功は、コピーされたファイルを確認して手動で対応する。
- ビルド後の通知も行わない
- PowerShellで実装する。
- ビルドの対象は、Releaseビルドのみとする。

# 対象のプロジェクトとパス
- TaskStarter
	- TaskStarter.slnx
	- コピー元: D:\git\PrivateTools\TaskStarter\TaskStarter\bin\Release
	- コピーするファイル
		- TaskStarter.exe
		- RelatedTool.bat
- MakePlayList
	- MakePlayList.slnx
	- コピー元: D:\git\PrivateTools\MakePlayList\MakePlayList\bin\Release
	- コピーするファイル
		- MakePlayList.exe
- PlayListToTaskStarterConverter
	- MakePlayList.slnx
	- コピー元: D:\git\PrivateTools\MakePlayList\PlayListToTaskStarterConverter\bin\Release
	- コピーするファイル
		- PlayListToTaskStarterConverter.exe

# コピー先
- \\OK-HOST-WIN11\share\common\TaskStarter

# 技術スタック
- PowerShell
- MSBuild (Visual Studioのビルドツール)
※上記ツールは、PowerShellスクリプト内で呼び出す形で使用する。

# 環境構築
- psファイルの拡張子紐づけ
	- PowerShellパス: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe


- MsBuildパス
- "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"

