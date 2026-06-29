
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



