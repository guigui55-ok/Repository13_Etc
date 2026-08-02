
# 実作業手順
260627

VM作成
ネットワーク構成：Default
VM名：VM4_Linux_Dev_DB.vhdx
ファイル場所：D:\VMs\
サイズ：127GB

Ubuntuを入手
Ubuntu Desktop 26.04 LTS
isoをVMにセット
VM起動
    メモリ不足エラー
VM1（Win_Dev）のメモリを12288 →8192に
Chromeを閉じる→OK
↓
UEFIで白画面、System was locked.～
↓
Hyper-V 第2世代VMの セキュアブート が原因で、UbuntuのISOを起動できていない

※フォルダを変更するときは
VMを右クリック移動で行う（設定画面ではない）
フォルダを変更： VM04_Linux_Db

↓
セキュアブートを無効にして起動

↓
iso起動
Try Install

Langugage：日本語
Ubuntuのアクセシビリティ なし
キーボード 日本語 OADG 109A
インターネット 有線
Ubuntuをインストール
インストール方式：対話式
アプリ: 既定の選択（最低限構成）
サードパーティソフトウェア：有効でOK
↓
動作が思ったより重い
サードパーティソフトウェアのインストール検討をきっかけに、OSの選別を再検討
↓
デスクトップ環境はほかのPC、VMで行うため、
Ubuntu Server LTS に変更
Ubuntu Server 26.04 LTS  ダウンロード
↓
言語： English
Keyboad
Layout: Japanese
Variant: Japanese OADG 109A
標準パッケージ（メモリ使用量はminimizedと大差ないため）
ネットワーク
何も変更せず、そのまま「Done
理由は、
    インターネットに接続できる
    apt update や Docker のインストールができる
    学習初期はネットワーク設定で詰まるリスクを減らせる

Proxy空欄
↓
Ubuntu archive mirro configuration で止まる
VM再起動

↓
進んだ
Disk設定 そのまま
File System

Favatdb

Ubuntu Pro→契約しない
必須: SSHインストール

VMが止まってしまうため、以下に変更（詳細な検討はしていない）
（おそらくメモリ不足と仮定）
最小メモリ　2048に変更
プロセッサ　2
メモリ設定　6144


※注意
インストール中に、LinuxVMから、他のウィンドウをアクティブにすると、VMが止まる（キーボード操作が不能になっているだけかもしれないが、メモリ不足が濃厚）

※最小メモリ512，1024だと、VMが止まってしまうことがあった。

-------------
動作が思ったより重かったため、Ubuntuより軽いものに移行することも考えておく
追記
基本的に余計なソフトウェア・コンポーネントは入れない

------------
【重要】
上記のフリーズ減少は、動的メモリ設定ONが原因の可能性あり。
動的メモリ：OFF にしておく。
メモリは4GBでOK


メモリ調査
Ubuntuから見えているメモリ量、実際にどれだけ使われているか
free -h


cat /proc/meminfo
　起動しているプログラムすべてのメモリ使用量を表示
　　長すぎて見切れる。

シャットダウンコマンド
sudo shutdown -h now
または、sudo poweroff

↓
起動OK、安定動作OK
↓


↓
OS初期設定
「ミドルウェア（DockerやSQL Server）を入れる前に、OSを運用できる状態に整える作業」

↓
VM3、LinuxDBの
・IPアドレス確認済み
・SSHサービス起動確認済み
・Ping確認
　ホストPCからのpingは通るが、VM2（WinDev）からのpingは通らない
　※VM2のMobaXTermから実行したい。

---------
↓
・VM2はDefault Switch
・VM2のIPアドレス：172.22.32.1
・VM3のIPアドレス：172.17.133.21

↓
Default SwitchをExternal Switchにする。

                    自宅ルータ
                        │
                192.168.1.0/24
                        │
                External Switch
                        │
      ┌─────────┼──────────┐
      │         │          │
 WindowsVM   UbuntuDB   ProxyVM
 192.168.1.x 192.168.1.x 192.168.1.x

↓
設計書への記載例

■ ネットワーク構成

Hyper-VのExternal Virtual Switchを使用する。

各VMは同一ネットワークに接続し、
以下を可能とする。

・VM間通信
・ホストPCからのSSH接続
・インターネット接続
・DockerおよびSQL Serverの通信確認

IPアドレスはDHCPによる自動取得とし、
必要に応じて固定IPへ変更する。

↓
構築手順書

1. External Virtual Switchを作成する
2. VMのNetwork AdapterをExternal Switchへ変更する
3. Ubuntuを再起動する
4. IPアドレスを確認する
5. Windows VMからPing確認
6. MobaXtermでSSH接続確認

↓
VM3→VM2 間ネットワークがつながらない

↓
Internal Switch（NIC）を増設

◆設計書に追記
ネットワーク構成方針

Hyper-V上にExternal SwitchおよびInternal Switchを構成する。

External Switchはインターネット接続およびホストPCからの管理（SSH等）に使用する。

Internal SwitchはVM間の内部通信専用ネットワークとして使用し、
DB通信や各サーバー間通信を外部ネットワークから分離する。

これにより、
・サーバー間通信のセキュリティ向上
・役割ごとのネットワーク分離
・将来的なWebサーバー、Proxyサーバー等の追加に対応しやすい構成
を実現する。

↓
Internal Switchを作成
VM2,3に追加
（Hyper-v 設定＞ハードウェアの追加＞ネットワーク）
Ubuntu確認
ip addr

↓
ホストPCの設定変更
Win＋R → ncpa.cpl

Internal Switch
IPアドレスを設定、サブネットマスク 255.255.255.0


◆ネットワーク構成

| 機器                        | IPアドレス       | サブネットマスク      | デフォルトゲートウェイ |
| ------------------------- | ------------ | ------------- | ----------- |
| ホストPC (Internal)          | 172.16.10.1  | 255.255.255.0 | 空欄          |
| Windows VM (Internal NIC) | 172.16.10.10 | 255.255.255.0 | 空欄          |
| Ubuntu (eth1)             | 172.16.10.20 | 255.255.255.0 | 設定しない       |

各VMのIPを変更、ping確認
↓
MobaXterm SSH接続 OK
（あとは、SSH接続NGになったときにだけVMコンソールを使用する）
　　自動ログ設定ON
↓
UbuntuServer初期設定

OS情報確認  OK
時刻・ホスト名確認  OK
パッケージ更新  OK
必要ツール導入　OK
Docker導入  OK
SQL Server導入  OK

↓
DB環境などは別Linuxに移しやすいように、別HDD（VHDX）にする。
VMにVHDXを追加
保存場所はVMの近く  D:\VMs\VM04_Linux_Db\Virtual Hard Disks
DbData.vhdx
容量    127GB


C:\ProgramData\Microsoft\Windows\Virtual Hard Disks

↓
Docker構築　済
↓
DokerComposeをgit管理に
    フォルダ作成
    フォルダ権限調整
    gitのリポジトリ用意
    アクセスtoken用意
        token名は、端末ごと
        期限はプライベート用なら1年くらいでOK
            実務では　検証30-60-本番90～1年とか
        AddPermission
            Contants    Read And Write
            
↓
Sql Serverを構築
git更新
↓
Dev Win にSSMSをインストール









-----------
上記作業を通して、設計書などに変更が必要となる。

各ドキュメントの更新をするため、上記作業ログをChatGPTで解析し、更新箇所を洗い出す。

LogとChat回答について以下の通り内容を確定、および訂正する。

* OSの変更について
当初Ubuntu Desktopを採用したが、実作業においてGUIが不要であることとリソース消費を考慮し、Ubuntu Serverへ変更した。
↓
このことについては、おそらくVMの動的メモリ設定が誤っていたため、発生した事項である可能性が高い。
ただ、この現象を受けて再度検討した結果、本件の目的としてCLIでも全く問題ないため、これを採用した。

* 1-1. Ubuntu DesktopからUbuntu Serverへの変更 項目について
要確認事項について、「Ubuntu Server LTS」と表記します。（実行時の最新版でよい認識です）

* 1-2. ホストPCのリソース制約とVMリソース設定 について
Chat回答通り、CPUは「要確認」とする。

* 2-1. ネットワーク方式の変更 について
SQL Server接続は一旦Internalにします。（学習用途のため）
ただ、ホストPCから接続したいです。その場合はExternalでしょうか？

* 2-2. DBデータディスク分離方式の採用
「DbData.vhdx、127GB」について、
メインのドライブとは別のドライブです（WindowsでいうDドライブにしています。ext4ですかね？）
/srv/docker 等へマウントをしています。（記憶にあり、ターミナルログが別途あります）
SQL Serverデータも本当にそのVHDXへ配置済みです。（記憶にあり、ターミナルログが別途あります、すぐには出せませんが必要なら持ってきます）

* 2-3. SSH中心の管理方式への変更
SSH接続元をホストPCだけにするか、開発用Windows VMも正式な管理端末とするか。
→ ホストPCからも開発用Windows VMからもアクセスしたいです。
ホストPCにクライアントアプリを置いてDBアクセスしたいため

* 3-1. VM詳細設定
Serure BootはOFFにしていますので、その旨反映したいです。

* 3-2. Ubuntu Serverインストール設定
Favatdb はパスワードを忘れないように、パスワードの一部を記載したものです。
セキュリティ上ドキュメントには反映しないようにしたいですね。

* 3-3. Internal Network固定IP
ご指摘の通りにします。

* 3-4. Docker Compose構成管理方式
確かこの作業はしましたが、ターミナルログが必要ですかね？　必要であればご指示ください。

* 4. 構築手順への影響
- E. Docker ComposeとGit
1年を正式仕様にしてはいけません。→これはなぜでしょうか？
本番運用でも運用の現場によっては1年とする場合もあると聞きました。（確か）
この点については、ちょっと確認・検討したいですね。

* 5. テスト設計・テスト手順への影響
ご指摘の通り、どのIP、どのNIC、どのネットワーク経路で接続することを期待するか、を明記しましょう。

* 6. 運用設計への影響
- 6-3. DB用VHDXの運用
DBデータも消えてよい学習環境なのか　→消えてもよいが、なるべくは消したくない以降です。  
OSのみ再構築してDBデータは残したいのか　→はい、こちらを採用します。

* 7. 運用手順・障害対応手順への影響
ご指摘の通り追加します。

* 8. 調査結果や技術選定理由として残すべき事項
それぞれ、ご指摘の通り記載したいです。

* 9. 一時的な試行錯誤・変更不要事項
インストール時のフリーズについて、前述の通り、動的メモリ設定ONが原因の可能性が高いです。現時点でこの線で話を進めたいです。  





# 作業ログ
（途中から記載）
要求仕様～テストまでドキュメントを作成。
作業実施
（テスト未実施）
各種ドキュメントを更新（実作業にて設計と差異があったため）
- basic_design.md 更新完了（v2.0）






