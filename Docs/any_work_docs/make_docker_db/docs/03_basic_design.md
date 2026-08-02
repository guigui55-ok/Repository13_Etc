# 基本設計書（Basic Design）

案件名:
学習用サーバー環境構築および設計書整備

版数:
v0.2

---

# 1. 目的

本書は、学習用サーバー環境の基本構成および設計方針を定義する。

対象環境は、Hyper-V上に構築するUbuntu Server LTSのVMとし、
Docker上でSQL Serverを利用可能な環境とする。

---

# 2. 前提条件

* ホストPCはWindows11とする
* 仮想化基盤はHyper-Vとする
* ゲストOSはUbuntu Server LTSとする
* DBはSQL Serverを採用する
* SQL ServerはDockerコンテナとして構築する
* DB管理クライアントはWindows側のSSMSを利用する
* 学習用途のため、本番運用は対象外とする

---

# 3. 全体構成

```text
ホストPC（Windows11）
  ├ VM1：Windows11（メイン開発用）
  ├ VM2：Linux（サブ開発用）
  ├ VM3：Linux（検証用）
  └ VM4：Ubuntu Server LTS（本件対象）
        ├ Docker Engine
        │    └ SQL Serverコンテナ
        └ データ用VHDX
             └ SQL Server永続データ・バックアップ領域
```

SSMSはWindows側に配置し、VM4上のSQL Serverコンテナへ接続する。

VM4は、外部通信に使用するネットワークと、ホストPCおよびWindows開発VMとの内部通信に使用するネットワークを分離する。

---

# 4. VM設計

## VM4基本設定

| 項目    | 設定                              |
| ----- | ------------------------------- |
| 仮想化基盤 | Hyper-V                         |
| OS    | Ubuntu Server LTS               |
| メモリ   | 4096MB（固定）                      |
| CPU   | 2 vCPU                          |
| OSディスク | OS用VHDX                         |
| データディスク | DbData.vhdx（127GB）                |
| 用途    | Docker、SQL Server、サーバー構築学習      |

## 設計方針

* サーバー用途としてCLIおよびSSHを中心に運用するため、Ubuntu Serverを採用する
* VMメモリは4096MBの固定割り当てとし、動的メモリは使用しない
* CPUは2 vCPUとする
* OS領域とDB関連のデータ領域を分離するため、データ用VHDXを使用する
* ホストPCのリソース制約を考慮し、VM同時起動数を制御する

---

# 5. OS設計

## 採用OS

Ubuntu Server LTS

## 採用理由

* 本環境の用途ではGUIを必要とせず、CLIおよびSSHによる管理で運用可能である
* サーバー用途として不要なGUI関連リソースを使用しない構成とできる
* 学習情報が多い
* DockerおよびSQL Serverの構築情報が多い
* 将来のサーバー構築・運用学習へ展開しやすい

---

# 6. ネットワーク設計

## 接続方針

* VM4は外部通信用と内部通信用の2系統のネットワークへ接続する
* 外部通信用ネットワークは、VM4からインターネットへの接続に使用する
* 内部通信用ネットワークは、ホストPCおよびWindows開発VMからVM4への管理通信・DB通信に使用する
* Ubuntu Serverの通常操作は、内部通信用ネットワークを経由したSSH接続を基本とする
* SQL Serverへは、内部通信用ネットワークを経由してSSMSから接続する
* インターネットへのサービス公開は行わない

## ネットワーク構成

| ネットワーク | 用途 | 方針 |
| ---------- | ---- | ---- |
| External Network | VM4からのインターネット接続 | OS・Docker等の更新および必要な外部通信に使用する |
| Internal Network | ホストPC・Windows開発VMとの内部通信 | SSHおよびSQL Server接続に使用する |

Internal Networkは `172.16.10.0/24` を使用し、以下のアドレス体系とする。

| 対象 | IPアドレス |
| ---- | ---------- |
| ホストPC | 172.16.10.1/24 |
| Windows開発VM | 172.16.10.10/24 |
| VM4（Ubuntu Server） | 172.16.10.20/24 |

Internal Network側にはデフォルトゲートウェイを設定しない。

## 公開範囲

| 対象         | 公開方針                         |
| ---------- | ---------------------------- |
| Ubuntu操作   | Internal Network経由のSSH接続を基本とする |
| SQL Server | ホストPCまたはWindows開発VMから接続         |
| Dockerコンテナ | 必要最小限のポートのみ公開                |
| インターネット公開  | 対象外                          |

---

# 7. Docker設計

## 利用方針

* Docker Engineを導入し、今後の開発・DB検証に利用可能な状態とする
* SQL ServerはDockerコンテナとして構築する
* コンテナ構成はDocker Composeで管理する
* Compose構成ファイルおよび関連する設定テンプレートはGitで管理する
* 認証情報などの秘密情報を含むファイルはGit管理対象外とする
* 秘密情報を除いた設定例はテンプレートとしてGit管理する
* 本フェーズでは複雑なコンテナオーケストレーションは対象外とする

---

# 8. DB設計

## 採用DB

SQL Server

## 管理ツール

SSMS

## 配置方針

* SQL ServerはVM4上のDockerコンテナとして構築する
* SQL Serverのデータおよびバックアップ領域は、コンテナ内部に保持せず、VM4上のデータ用VHDXへ配置する
* コンテナからホスト側の永続化領域をbind mountして使用する
* SQL Serverコンテナを再作成した場合でも、DBデータを継続利用できる構成とする

## データ配置方針

| 用途 | VM4上の配置先 |
| ---- | ------------ |
| SQL Serverデータ | `/srv/docker/sqlserver/data` |
| SQL Serverバックアップ | `/srv/docker/sqlserver/backup` |

データ用VHDXはVM4上で `/srv/docker` にマウントし、SQL Serverの永続データおよび関連するDB資産をOS領域から分離する。

## 接続方針

* SSMSはWindows側にインストールする
* ホストPCおよびWindows開発VMから、Internal Networkを経由してSQL Serverへ接続できる構成とする

---

# 9. セキュリティ設計

## 基本方針

学習用途の閉域環境として構築し、インターネットへのサービス公開は行わない。

## セキュリティ方針

* 不要なサービスは起動しない
* 不要なポートは開放しない
* 初期パスワードは変更する
* 認証情報はドキュメントへ平文記載しない
* 認証情報などの秘密情報を含むファイルはGit管理しない
* 管理者権限の常時利用は避ける
* OS・Docker・SQL Serverは構築時点で更新する

---

# 10. 運用設計方針

詳細は運用設計書で定義する。

本基本設計では以下の方針のみ定義する。

* VM起動・停止手順を整理する
* 構築手順書により再構築可能とする
* SQL Serverの永続データはデータ用VHDX上で管理する
* Docker Composeによるコンテナ起動・停止・状態確認方法を整理する
* 障害時は、構成ファイルと永続データを利用した再構築を基本とする
* スナップショット取得を必要に応じて利用する
* 構成変更時はGit管理対象ファイルおよび関連ドキュメントの整合を確認する

---

# 11. テスト方針

詳細はテスト仕様書で定義する。

基本設計時点では以下を確認対象とする。

* VM起動確認
* Ubuntu Serverログイン確認
* SSH接続確認
* Docker動作確認
* Docker ComposeによるSQL Serverコンテナ起動確認
* Internal Network経由のSSMS接続確認
* データ用VHDXの利用確認
* SQL Serverデータ永続化確認
* Git管理対象・対象外の確認
* ドキュメント整備確認

---

# 12. 制約事項

* ホストPCのメモリは32GBである
* VM1、VM2、VM3、VM4の同時起動は避ける
* 本環境は学習用途であり、本番運用・高可用性構成は対象外とする
* Oracle、PostgreSQL等の追加DB構築は本フェーズ対象外とする
* Docker Swarm、Kubernetes等のコンテナオーケストレーションは本フェーズ対象外とする
