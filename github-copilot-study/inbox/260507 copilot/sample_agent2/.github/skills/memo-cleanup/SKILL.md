---
name: memo-cleanup
description: "自由形式メモを構造化 Markdown に変換する手順"
---

# Memo Cleanup Skill

## 実行手順

1. `work/memo.txt` を読む
2. 情報を抽出する
3. 類似内容をグループ化する
4. 見出し構造を決定する
5. Markdown に変換する
6. `work/output.md` に出力する

## 出力形式

```markdown
# [トピック]

## [分類]
- 項目

## [次アクション]
- TODO
```

## 完了条件

* 情報欠落なし
* 重複整理済み
* Markdown 構造化済み

## トリガー例

* work/memo.txt を整理してください
* output.md に反映してください
* メモ整理を実行してください

## 出力ルール

スキル利用時は先頭に出力する。

[used skill memo_cleanup]

変更したファイル一覧を出力する。