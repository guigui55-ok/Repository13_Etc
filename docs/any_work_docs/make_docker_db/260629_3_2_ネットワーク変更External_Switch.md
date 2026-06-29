手順は以下です。

## 1. External Switchを作成

1. Hyper-V マネージャーを開く
2. 右側の **仮想スイッチ マネージャー**
3. **外部** を選択
4. **仮想スイッチの作成**
5. 名前を付ける
   例：`External_LAN`
6. 接続の種類で **外部ネットワーク** を選択
7. 物理LANアダプターを選択
   有線LANなら有線を推奨
8. **管理オペレーティングシステムにこのネットワークアダプターの共有を許可する** にチェック
9. OK

※ 一瞬ホストPCのネットワークが切れることがあります。

## 2. VMの接続先を変更

VM2、VM3を停止してから、

1. VMを右クリック
2. **設定**
3. **ネットワーク アダプター**
4. 仮想スイッチを `Default Switch` から `External_LAN` に変更
5. OK

VM2、VM3の両方で実施します。

## 3. 起動後にIP確認

Ubuntu側：

```bash
ip addr
```

Windows VM側：

```cmd
ipconfig
```

両方が同じネットワーク帯、例：

```text
192.168.1.xxx
```

になっていればOKです。

## 4. 疎通確認

VM2からUbuntuへ：

```cmd
ping UbuntuのIP
```

MobaXtermで：

```text
Host: UbuntuのIP
User: tok
Port: 22
```

これでSSH接続できれば成功です。

設計書には「Hyper-V External Switchを使用し、VM間通信・ホストPCからの接続・インターネット接続を可能とする」と残せば十分です。



VLAN IDが設定されていると、有線LAN、ネットワークがつながらなくなる。
「管理オペレーティングシステムで仮想LAN IDを有効にする」はチェックOFF
