# コンセプト
PlayListMakerの設定ファイルから、TaskStarterの設定ファイルに変換するツール

## 目的
PlayListMakerは、動画編集のためのプレイリストを作成するツール。
一方、TaskStarterは、複数のプログラムを一括実行するためのツール。
本ツールは、PlayListMakerの設定ファイルからTaskStarterの設定ファイルに変換することで、
TaskStarter（または、PlayListMaker）の設定ファイルの作成の省略を目的とする。

## 仕様
- 入力ファイルは、PlayListMakerの設定ファイル（INI形式）とする。
- PlayListMakerの設定ファイルを入力とする。
- 入力ファイル、コマンドライン引数で指定する。
- 出力ファイルは、コマンドライン引数で指定したフォルダパスに入力ファイル名を合わせたパスとする。
- 出力は、TaskStarterの設定ファイル（INI形式）とする。

### 入力ファイル例
本項目を仕様とする。

#### PlayListMakerの設定ファイル例
[Playlist1]
name=SamplePlaylist1
inputFolder1=C:\Videos
outputFolder=C:\Playlists
extensions=.mp4,.avi,.mkv
includeSubfolders=true
randomize=false
FilePath=C:\ZMyFolder\document\TaskStarter
FileName=Laa2.xspf
Parameter=""
Position=50,0
WindowRect=680,520

[PlayList2]
... 以下同様に続く

#### TaskStarterの設定ファイル例
[Task1]
name=SamplePlaylist1
FilePath=C:\ZMyFolder\document\TaskStarter
FileName=Laa2.xspf
Parameter=""
Position=50,0
WindowRect=680,520
includeSubfolders=false
Random=true
inputFolder1=D:\Movies
inputFolder2=D:\Movies2

[Task2]
... 以下同様に続く

### 変換ルール
- 基本的にPlayListMakerとTaskStarterの設定項目はそれぞれ異なるため、セクション以下の項目は変更しない。
- セクション名の変更のみを行う。
    - 以下の文字列を置換する。(正規表現）
        - 他の項目内に同じ文字列が含まれている可能性があるため、セクション名のみに適用すること。
        - 正規表現マッチ例
            - ^\[Playlist(\d{1,2})\]$ -> [Task$1]

## 動作確認・テスト用メモ
ディレクトリ移動
cd /d D:\git\PrivateTools\MakePlayList\PlayListToTaskStarterConverter\bin\Debug
入力ファイル
D:\git\PrivateTools\MakePlayList\PlayListToTaskStarterConverter\TestFiles\testSetting1.txt
出力フォルダ
D:\git\PrivateTools\MakePlayList\PlayListToTaskStarterConverter\TestFiles\Output

PlayListToTaskStarterConverter.exe "D:\git\PrivateTools\MakePlayList\PlayListToTaskStarterConverter\TestFiles\testSetting1.txt" "D:\git\PrivateTools\MakePlayList\PlayListToTaskStarterConverter\TestFiles\Output"


# TODO
- 置換ルールの追加
    - // 行頭が "outputFolder=" の場合、"FilePath=" に置換する
    - 上記は実装済みなのでドキュメントに反映すること。 