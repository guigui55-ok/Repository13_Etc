
# 事前調査
本件の目的は、concept.md を参照。　 

# 背景
詳細はcondept.mdを参照　 
サーバー構築経験習得のため　 

# 補足
GUIでも操作したいため、デスクトップ版Ubuntuを採用。　 

## VMの使用環境
本件で構築するVM、DBは、以下の環境で使用する。　 
　 
ホストPC - Windows11　 
    Spec:　 
        CPU名: Intel(R) Core(TM) i7-6700K CPU @ 4.00GHz　 
        プロセッサ数: 8　 
        メモリ: DDR4  32.0GB　 
　L VM1(Hyper-V) - Windows11 (メイン開発用)　 
    Spec: 　 
        メモリ: 12288MB　 
        プロセッサ: 4　 
　L VM2(Hyper-V) - Linux (サブ開発用、Linuxの汎用的な作業用、今後Windowsから乗り換える可能性あり)　 
    Spec: 　 
        メモリ: 8192MB　 
        プロセッサ: 4　 
　L VM3(Hyper-V) - Linux (サブVM、Linuxの汎用的な作業用、今後Windowsから乗り換える可能性あり)　 
　　※あまり使用しない想定　 
　　　HostPCの代わりにできるかを検証するためのVM　 
    Spec: 　 
        メモリ: 4096MB　 
        プロセッサ: 4　 
　L VM4(Hyper-V) - Linux（★本件で作成するVM）　 
    Spec: 　 
        メモリ: 8192MB　 
        プロセッサ: 4　 
    本VMのスペックの詳細は、後述にて記載。
　 
※HDDは潤沢にあるため、考慮不要。　 

# 懸念事項
* ホストPCは 32GB RAMで、既存VM1が12GB、VM2が8GB、VM3が4GB想定なので、全部同時起動するとすでに24GB使います。そこにVM4を8GBで足すと、ホストOS分を含めてかなり厳しいです。  
  
実務想定の判断としては、
VM4を使う日は、VM3は停止。できればVM2も停止。VM4は 8GB / 4コアで開始。重ければ12GBへ増やす。  


## VM4のスペック方針

初期構成は以下とする。  

- メモリ: 8192MB
- プロセッサ: 4
- ディスク: 100GB以上（HDDは潤沢なため余裕を持たせる）
- OS: Ubuntu Desktop LTS
- 用途: Docker、SQL Server、今後のDB検証、サーバー構築学習

理由:  
Ubuntu DesktopをGUI操作で利用し、DockerおよびSQL Serverを動作させるため、最低構成ではなく実用構成として8GB/4コアを採用する。  

制約:  
ホストPCのメモリは32GBのため、VM1/VM2/VM3/VM4を同時に全て起動する運用は避ける。  
VM4使用時は、VM3を停止し、必要に応じてVM2も停止する。  

見直し条件:  
- GUI操作が重い
- Dockerコンテナ起動時にスワップが多発する
- SQL Server操作が遅い
- ホストPC側が重くなる

上記の場合は、VM4のメモリを12GBへ増やすか、同時起動VMを減らす。  


* スペック決定理由
- Ubuntu Desktop + Docker + SQL Server想定で、GUI操作もしたい構成のため。  
- Docker公式は Ubuntu 26.04 LTS / 24.04 LTS をサポート対象
- SQL Server on Linux は最低 2GB RAM・2コアだが、実用上の余裕を見る。






# ------------------
# 初期検討内容
## 候補比較
- Hyper-V
- VirtualBox
- Docker Desktop
- WSL2
↓
現在Hyper-V使用しており、メモリ等の資源の節約のためHyper-Vを採用。

## OS比較
- Ubuntu
- Rocky Linux
- Debian
↓
Ubuntu Desktop LTS を採用する。
↓
理由:
本件は「サーバー実務学習」が目的だが、サーバー構築経験が少ないため、
まずはGUI操作可能な環境で構築・設定・運用経験を得ることを優先する。

また、Docker・SQL Server・将来的なDB追加・個人開発環境としても流用可能であり、
学習コストと拡張性のバランスが最も良いため採用する。

将来的には同構成を Ubuntu Server または Rocky Linux に移植し、
CLI主体運用も学習する。

## DB比較
- PostgreSQL
- SQL Server
- MySQL
↓
本業にて、直近で使用しているのが SQL Serverのため、これを採用。
その他も構築予定だが、今回は作業をし無い。（今後構築することを想定して作業を進める）

## 構成比較
上記の通り目的が明確なため、別案は用意しない。


## Hyper-VのVM設定比較
| 用途   |     メモリ |     CPU | コメント                                     |
| ---- | ------: | ------: | ---------------------------------------- |
| 最低限  |     6GB |     2コア | Ubuntu Desktopだけなら可。ただしDB/Dockerで重くなりやすい |
| 推奨   | **8GB** | **4コア** | 今回のVM4の初期値として妥当                          |
| 余裕あり |    12GB |     4コア | VM1/VM2を同時起動しない日なら快適                     |

ただしホストPCは 32GB RAMで、既存VM1が12GB、VM2が8GB、VM3が4GB想定なので、全部同時起動するとすでに24GB使います。そこにVM4を8GBで足すと、ホストOS分を含めてかなり厳しいです。

実務想定の判断としては、
VM4を使う日は、VM3は停止。できればVM2も停止。VM4は 8GB / 4コアで開始。重ければ12GBへ増やす。




# ------------------
# 不明点
VM4はどのくらいのスペックにすべきか？（メモリ、プロセッサ数）
→解決済み

