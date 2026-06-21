# 基本設計（Architecture）
目的：
「どう作るか」

Windows11
└─ Hyper-V
   └─ VM3 Ubuntu
      ├─ Docker
      │  ├─ nginx
      │  ├─ PostgreSQL
      │  └─ sample-app
      │
      └─ Backup

ここで決めるもの：

ネットワーク
VM分割
Docker方針
ストレージ
セキュリティ

