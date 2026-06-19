---
name: memo-prioritizer-skill
description: "自由形式メモ（Todoリスト）に優先順位をつけて、構造化 Markdown に変換する手順"
user-invocable: true
---
## 実行手順

1. `work/memo.txt` を読む
2. 情報を抽出する
3. 類似内容に優先順位をつける
4. Markdown に変換する
5. `work/output.md` に出力する

## トリガーキーワード
以下のキーワードをトリガーとする。
* memo.txt , TODO , 整理

## 出力ルール
- スキル利用時は、`[used skill memo-prioritizer-skill]`を出力ファイルの先頭に出力する。
- 余分な空白は削除する。
- 呼び出し元のagent.mdがわかる場合は、`[used agent {aggent name}]`をファイルの先頭(used skill の前)に出力する。（used skill を消さない）

## 補足
- 優先順位を付けた理由について、ファイル下部に別途記載する。
- `**/achive/**`フォルダ内のファイルは読み取らない