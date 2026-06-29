これは**今だからこそやっておく価値があります**。

現在は Docker の保存先が

```text
/srv/docker/docker-data
```

になっています。

これを

```text
/srv/docker-data
```

へ移動し、Dockerの設定も変更します。

---

# 手順1 Docker停止

```bash
sudo systemctl stop docker
```

確認

```bash
sudo systemctl status docker --no-pager
```

---

# 手順2 ディレクトリ作成

```bash
sudo mkdir -p /srv/docker-data
```

---

# 手順3 データ移動

`mv` よりも `rsync` をおすすめします。

```bash
sudo rsync -aHAX /srv/docker/docker-data/ /srv/docker-data/
```

コピー後、

```bash
ls /srv/docker-data
```

で内容を確認します。

---

# 手順4 daemon.json変更

```bash
sudo nano /etc/docker/daemon.json
```

現在

```json
{
    "data-root": "/srv/docker/docker-data"
}
```

となっているので、

```json
{
    "data-root": "/srv/docker-data"
}
```

へ変更します。

---

# 手順5 Docker起動

```bash
sudo systemctl start docker
```

---

# 手順6 確認

```bash
docker info
```

確認項目

```text
Docker Root Dir: /srv/docker-data
```

---

# 手順7 動作確認

```bash
docker ps
```

```bash
docker images
```

今までのコンテナ・イメージが見えれば成功です。

---

# 手順8 古いフォルダ削除

問題ないことを確認したら

```bash
sudo rm -rf /srv/docker/docker-data
```

---

# 完成形

```text
/srv
├── docker
│   ├── .git
│   ├── .gitignore
│   ├── sqlserver
│   ├── nginx
│   └── redis
│
└── docker-data
```

---

## なぜこちらが良いのか

ここが実務で重要な考え方です。

### `/srv/docker`

これは**構成管理（Infrastructure as Code）**です。

Gitで管理するものだけ置きます。

```text
compose.yaml
.env.example
README.md
```

---

### `/srv/docker-data`

これは**実行環境**です。

Dockerが自動生成する

* イメージ
* コンテナ
* ネットワーク
* ボリューム
* BuildKit

などが保存されます。

Gitで管理するものではありません。

---

## 私ならさらに一歩進めます

今回の構成なら、Dockerのデータ保存先もSQL Serverのデータ保存先も、**データ用VHDX上**にあるので、OSと完全に分離できます。

例えばマウントポイントを次のように整理すると分かりやすいです。

```text
/srv
├── docker        ← Git管理（構成・設定）
└── data
    ├── docker    ← Docker Root Dir
    ├── sqlserver
    ├── backup
    └── logs
```

あるいは

```text
/srv
├── docker
└── docker-data
```

でも十分です。

どちらも実務で見かける構成ですが、**「Gitで管理するもの」と「実行時に生成されるもの」を物理的に分ける**という考え方が最も重要です。
