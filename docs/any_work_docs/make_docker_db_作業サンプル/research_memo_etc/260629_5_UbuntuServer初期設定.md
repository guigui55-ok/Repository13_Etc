**「OS初期設定」**

いま完了したことは、

* Ubuntu Server構築
* メモリ安定化
* Internalネットワーク設定
* MobaXtermでSSH接続

なので、次はSSH上で以下を進めます。

```bash
cat /etc/os-release
hostnamectl
free -h
df -h
timedatectl
```

その後、

```bash
sudo apt update
sudo apt upgrade
```

です。

流れとしては、

1. OS情報確認
2. 時刻・ホスト名確認
3. パッケージ更新
4. 必要ツール導入
5. Docker導入
6. SQL Server導入

です。まずはSSHで上の確認コマンドからでよいです。
