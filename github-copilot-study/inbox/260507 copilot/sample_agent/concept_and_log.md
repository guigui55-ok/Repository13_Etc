# 本ワークスペースの目的と概要
github Copilot の agent.md について学習したいと考えています。 
（AGENT.md の記載方法・格納場所を学習する）
まず、基本機能を理解するために、copilot-instructions.md , AGENT.md のみの最小構成で実行したいと考えています。 
題材は、メモの整理で、以下のようなメモを整理をしたいです。 （メモは本ファイル最下部に記載）
メモは work/memo.txt に記載しておき、それを処理して work/output.md に出力する。

AGENT.md が使用されたときは、その旨をファイルまたはCopilot回答欄に出力する。

AGENT.md は今後複数作成する予定（同じワークスペースに同居させる）なので、これも意識したい。

# フォルダ構成（最小）
workspace/
├─ .github/
│  ├─ copilot-instructions.md
│  └─ AGENT.md
├─ memo.txt
└─ output.md


# サンプルメモ内容

```
【旅行メモ】

行先：
小樽

候補：
・運河を見る
・海鮮丼を食べる
・ガラス工房を見る

条件：
・日帰り
・予算1万円以内
・朝はゆっくり出発したい

気になっていること：
・電車がよいか車がよいか
・混雑しそうか
・雨の場合どうするか
```


# チャット指示
memo.txt を整理して output.md に反映してください。

AGENT.md に従って memo.txt を整理してください。
結果は output.md に保存してください。

# Ask/Agent実行比較
* Askで実行
work/memo.txt を整理して output.md の内容案を出してください。
memo-organizer を使ってください。
→大した変わらない


## まとめ
Agent.md は .github/agents フォルダを作成して、その中に `[summay].agent.md' ファイルを作成する。
→VsCodeチャット欄の作業を選ぶ場所（Agent/askなど）に表示されるので、それを選択して質問をする。
