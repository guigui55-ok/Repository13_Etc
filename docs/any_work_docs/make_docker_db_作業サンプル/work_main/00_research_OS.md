# 事前調査 参考資料 OS比較のみ

## OS比較

### 比較候補
- Ubuntu Desktop LTS
- Rocky Linux
- Debian

### 比較観点

|観点|Ubuntu LTS|Rocky Linux|Debian|
|---|---|---|---|
|学習容易性|◎|○|△|
|情報量（日本語）|◎|○|○|
|Docker利用|◎|◎|◎|
|GUI利用|◎|○|△|
|DB構築|◎|◎|○|
|実務サーバー理解|○|◎|○|
|長期保守|◎|◎|◎|
|個人開発との相性|◎|○|○|

### 候補ごとの評価

#### Ubuntu Desktop LTS
メリット
- 学習資料・トラブルシュート情報が多い
- GUI・CLI両方で操作可能
- Docker、SQL Server、PostgreSQL等の情報が豊富
- 個人開発～サーバー学習まで幅広く利用可能
- LTSにより長期サポートがあり安定運用可能

デメリット
- 実務サーバーではGUI無し構成も多い
- Enterprise系Linuxより管理文化がやや異なる

#### Rocky Linux
メリット
- 実務サーバーに近い運用を学びやすい
- RHEL系知識が得られる
- サーバー用途に強い

デメリット
- GUI用途ではUbuntuより扱いにくい
- 個人学習情報が少し減る

#### Debian
メリット
- 軽量・安定
- Linux理解が深まる

デメリット
- 初学者向け情報はUbuntuより少ない
- GUI操作前提なら恩恵が少ない

### 採用方針

Ubuntu Desktop LTS を採用する。

理由:
本件は「サーバー実務学習」が目的だが、サーバー構築経験が少ないため、
まずはGUI操作可能な環境で構築・設定・運用経験を得ることを優先する。

また、Docker・SQL Server・将来的なDB追加・個人開発環境としても流用可能であり、
学習コストと拡張性のバランスが最も良いため採用する。

将来的には同構成を Ubuntu Server または Rocky Linux に移植し、
CLI主体運用も学習する。

## 今後の展望
Phase1：Ubuntu Desktop（今回）
Phase2：Ubuntu Server（GUIなし移植）
Phase3：Rocky Linux（RHEL系比較）