以下の手順で進めてください。
前提：GitHub側に `ubuntu-docker-learning` を作成済み。

## 1. `/srv/docker` に移動

```bash
cd /srv/docker
```

## 2. root所有になっていないか確認

`sudo git init` しているので確認します。

```bash
ls -ld .git
```

`root root` になっていたら修正します。

```bash
sudo chown -R tok:tok .git
```

## 3. Gitユーザー設定

未設定なら設定します。

```bash
git config --global user.name "Toshi Ok"
git config --global user.email "GitHubに登録しているメールアドレス"
```

確認：

```bash
git config --global user.name
git config --global user.email
```

## 4. mainブランチに変更

```bash
git branch -m main
```

## 5. `.gitignore` 作成

```bash
nano .gitignore
```

内容：

```gitignore
# secrets
.env
*.key
*.pem
*.pfx

# logs
*.log

# database files
*.bak
*.mdf
*.ldf

# temporary
tmp/
```

保存：`Ctrl + O` → Enter → `Ctrl + X`

## 6. 状態確認

```bash
git status
```

## 7. 最初のコミット

```bash
git add .
git commit -m "Initial Docker learning environment"
```

## 8. GitHubリポジトリを登録

GitHubのURLに置き換えてください。

```bash
git remote add origin https://github.com/ユーザー名/ubuntu-docker-learning.git
```

確認：

```bash
git remote -v
```

## 9. GitHubへPush

```bash
git push -u origin main
```

これで `/srv/docker` 配下が GitHub の `ubuntu-docker-learning` に管理されます。
