# 目的
DBやテーブルの初期作成・構築について学習する。

# 次フェーズ
本ワークスペースは、そのための調査を行い、それに基づき実施する。  
また、実施するためのファイルなども、格納しておく。（gitに登録）  

# 背景
DB・サーバーの学習のためで、現状はサーバーとDBの構築が完了したため、そのあとのDB・テーブル作成とそのデータアクセスをする。  
また、別の環境（ほかのOSやDB種に変わった時もなるべく流用などできるようにしておきたい）

# 備考
実務的な作業を意識したい。  
今後同じようなことがあった場合、流用できればよい。  
現場によっては実務的に同じ環境を多数作成することがあると思うが、このDB環境を作る際にはそれをいくつか自動化しているはず。  
その作業を想定する。  
まず、具体的な手段を調査してから、ある程度設計したうえで実行する。

# 対象
* 環境： 別PCのDBサーバー内のDockerコンテナ上のDB  
* 別PCはHyper-VのVM
* 上記対象のVM： Ubuntu 26.04 LTS
* DBサーバーとSQLのバージョン： 
Microsoft SQL Server 2022 (RTM-CU25) (KB5081477) - 16.0.4255.1 (X64)
        Apr 23 2026 22:38:54
        Copyright (C) 2022 Microsoft Corporation
        Developer Edition (64-bit) on Linux (Ubuntu 22.04.5 LTS) <X64>                
  
* クライアント： Windows11 25H2
* クライアントにはSSMSをインストール済み


# 作成するデータ
テスト用のテーブル２つくらい。  
テストデータも登録しておきたい。（このデータ作成も自動化する）  

# 調査内容1
* DBやテーブル初期構築・作成の、実務的・効率的・一般的手法などについてどのような手法があるかを調査。（今回は採用した１案を実施する）

* DBスキーマのバージョン管理
    SQLファイルをどのように管理するか
    変更履歴をどう残すか
    FlywayやLiquibaseのようなマイグレーションツールを採用するか
    →簡単に概要だけ

* 複数環境への展開方法
    開発・検証・本番で同じSQLをどう適用するか
    パラメータ（DB名や接続先）の切り替え方法
    Docker環境以外にも流用できる構成にするにはどうするか

# 目標
docker compose up
DBが自動生成
テーブル作成（./build-db.sh）
テスト/サンプルデータ投入
即利用可能

# 作業の流れ
Phase1  調査
Phase2  方式決定
Phase3  設計
Phase4  実装
Phase5  手順書作成


# 調査内容
★必須
    実務ではどう作るか
    SQL Serverではどう作るか
    Dockerではどう作るか
★知っておきたい
    sqlcmd
    Docker Init
    Python
★概要だけ
    Flyway
    Liquibase
    スキーマ管理
    複数環境展開
    他のDBでの作成方法（Oracle、PostgreSQL）


-----


# 初期検討中
現在以下の方法を検討している。



# 調査の成果物
01_DB作成方式調査.md
02_テーブル作成方式調査.md
03_テストデータ投入方式.md
04_初期構築自動化.md
05_DBバージョン管理.md
06_運用更新方法.md
07_方式比較.md
08_採用方式.md



# 検討2
上記内容をもとに、以下の内容を検討中。

* DB作成方法の採用候補
    SQL + sqlcmd　★
    Docker Init SQL（概要や具体主要・手順を少しだけ理解する）
    python（概要や具体手法・手順を少しだけ理解する）

* 標準構成を設計
- フォルダ・ファイル構成例
db/
├── create_database.sql
├── create_tables.sql
├── create_indexes.sql
├── create_views.sql
├── create_proc.sql
├── seed_data.sql
└── build.sh
- 構成例2
db/
    001_database.sql
    010_tables.sql
    020_master.sql
    030_testdata.sql
