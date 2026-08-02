## 概要
C#.net のGetTempFileNameを多量に繰り返し実行したときに、重複するファイル名が出力されるかを検証する

# 方針

## 詳細
- 1つ以上のスレッドで実行可能にする
- 複数のプロセスでも実行可能にする
- １つのプロセスに１つ以上のスレッドを持ち、１スレッドにつき１つのワーカー（＝ループ）を持つ
- １つのループ内でGetTempFileNameを繰り返し実行する。
- ループの実行時間の間隔を変更可能にする
- ループ内で生成したファイル名は、ワーカー内にListで保持する。
- プロセス数、１プロセス当たりのスレッド数、１スレッドあたりの反復回数、１スレッドあたりの１回実行当たりの時間（ウェイト）は個別に設定可能にする
- 上記各設定値は引数から指定する。
- デバッグモードではプログラム内（Main冒頭）で設定値を変更できる。
- １つのワーカーの処理が終わると、ListをHashSet<string>にも保持させる。
- ListとHashSetの個数を比較して、重複しているかを確認する。
- Listの中をチェックして重複しているかチェックする。
    - 重複している場合は別途データを保持する。
- Listの内容はLogに出力する。
    - ログの内容は PID、ThreadID、生成ファイル名
      （どのPID、ThreadIDでどのファイル名を生成したかを確認したいため）
    - ただ、重くなる可能性があるのでログ出力モード（詳細、シンプル）で切り替えられるようにする。
      - 詳細：PID,ThreadID,生成ファイル名とサマリすべて出力
      - 簡易：Logには件数、重複有無、エラー、開始終了などのサマリを出力する。
- ログの他にも生成ファイル名保存専用の結果ファイルにも出力する。
    - こちらは生成ファイル名のみを記録する。（フルパスではない）（生成ファイル名が競合するかのチェックのため）
- 結果ファイルは、１プロセスごとにまとめて出力する。
- 最終的に、結果ファイルは、プロセスごとの結果リストも１つにまとめる。
- ログや結果ファイル出力のタイミングは、1プロセスが終わった段階とする（ループ途中でファイル出力はしない）
- できればログファイルは最後に1つにまとめたい（1回実行＝1つのログファイル、1つの結果ファイル）（中間ファイルは取っておいてよい）
- Logクラスは既存実装済みのAppLoggerを使用。
    - Logの出力はConsole、Debug、ファイル出力の3つにする。
- 最後に生成したtmpファイルはすべて削除する。
- エラーが発生したときはLogに出力して継続する。
- 各workerは中間ログ/結果ファイルを出力する
- 生成ファイルの出力先フォルダパスは1つのみを指定する想定。
  - そのため結果ファイルリストにはファイル名のみ記録する。
  - 複数フォルダを指定して、削除漏れした場合は手動で削除する。
- 重複したファイルがあった時、さらに別のファイル`duplicates.txt`に出力する
- 結果ファイルに加えメタ情報ファイルも作成しておく。 `merged_detail.csv`
  - 内容例：
    FileName,FullPath,ProcessId,ThreadId,WorkerIndex
    tmp1234.tmp,C:\Temp\xxx\tmp1234.tmp,12345,8,1


## 重複チェックについて
今回の目的は複数プロセス・複数スレッドでの重複確認なので、以下の3段階が必要。
① ThreadWorker内の重複チェック
② ProcessWorker内の重複チェック
③ merged_result.txt作成後の全プロセス横断チェック

## クラス、プロジェクト構成とプログラム実行モード
- 1つのプロジェクト内で完結する。
- 異なるプロセス間をまたぐ処理のため、複数モードを用意して、
　本exeを複数回実行する。

### モード
  - masterモード
      複数のworkerプロセスを起動する。
      すべてのworker終了後、結果ファイルとログファイルを集約する。

  - workerモード
      1つのプロセスワーカーを実行する。
      プロセスワーカー内で複数のスレッドワーカーを実行する。

### 構成
- メインproj
    - プロセスワーカークラス
    - スレッドワーカークラス
    - Logger
    - メイン実行クラス
        （複数のプロセスワーカーを実行＝exe呼び出し）
    - プロセス実行クラス
        （プロセスワーカークラスを実行、クラス呼び出し）

## フロー
- メインの実行（masterモード）
    - 自exeをworkerモードで1つ～複数回実行する
    - 上記プロセスがすべて終わるまで待つ
- 上記各プロセスのログをマージする
  - 一時ファイルをまとめて、1つのログファイル、結果ファイルを作成

# 詳細

## 出力ファイル構成
logs/
  run_20260515_231000/
    worker_001_12345.log
    worker_001_12345_result.txt
    worker_002_23456.log
    worker_002_23456_result.txt
    merged.log
    merged_result.txt
    merged_detail.csv
    duplicates.txt
    summary.txt

## クラス構成
TempFileNameTester
├─ Program.cs
├─ AppConfig.cs
├─ CommandLineParser.cs
├─ MasterRunner.cs
├─ WorkerProcessLauncher.cs
├─ ProcessWorker.cs
├─ ThreadWorker.cs
├─ TempFileRecord.cs
├─ WorkerResult.cs
├─ ResultFileWriter.cs
├─ ResultMerger.cs
├─ DuplicateChecker.cs
└─ AppLogger.cs

### 役割
Program
  引数を見て master / worker を切り替える

MasterRunner
  workerプロセスを複数起動
  終了待ち
  結果ファイル集約

ProcessWorker
  スレッドワーカーを複数起動
  プロセス単位の結果出力

ThreadWorker
  GetTempFileName を指定回数実行

DuplicateChecker
  List と HashSet で重複確認

ResultMerger
  workerごとの結果ファイルを1つにまとめる

## 引数例
TempFileNameTester.exe ^
  --mode master ^
  --process-count 4 ^
  --thread-count 8 ^
  --loop-count 10000 ^
  --wait-ms 0 ^
  --work-dir "C:\Temp\TestResult" ^
  --temp-dir "C:\Temp\GetTempFileNameTest" ^
  --delete-created-files false

masterがworkerを起動するときは、同じexeをこう呼びます。

TempFileNameTester.exe ^
  --mode worker ^
  --worker-index 1 ^
  --run-id 20260515_231000 ^
  --thread-count 8 ^
  --loop-count 10000 ^
  --wait-ms 0 ^
  --work-dir "C:\Temp\TestResult" ^
  --temp-dir "C:\Temp\GetTempFileNameTest"

## temp-dir の扱いの注意

Path.GetTempFileName() は、通常は OS の一時フォルダを使います。
temp-dir を指定したい場合は、workerプロセス起動時に環境変数を設定する方針（以下※1参照）にすることが推奨されるが、
一旦、本設定は実行しない。

※1
TEMP=C:\Temp\GetTempFileNameTest
TMP=C:\Temp\GetTempFileNameTest

# 動作確認
まず以下で確認するのが安全です。

TempFileNameTester.exe ^
  --mode master ^
  --process-count 2 ^
  --thread-count 2 ^
  --loop-count 100 ^
  --wait-ms 0 ^
  --work-dir "C:\Temp\TestResult" ^
  --temp-dir "C:\Temp\GetTempFileNameTest" ^
  --delete-created-files true ^
  --log-mode simple

  次に負荷を上げます。

--process-count 4
--thread-count 8
--loop-count 10000

※注意点として、現状の実装では temp-dir はまだ Path.GetTempFileName() の生成先には反映していません。
必要になったら WorkerProcessLauncher 内の TEMP/TMP 上書き部分を有効化してください。