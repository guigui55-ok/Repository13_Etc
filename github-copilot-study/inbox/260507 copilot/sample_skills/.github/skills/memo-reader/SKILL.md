---
name: memo-reader
description: workフォルダ内のメモを読み、整理・更新するときに使用する
allowed-tools: shell
---

# Memo Reader Skill

このスキルは、
「メモ」
「整理」
「workフォルダ」
「output.md」
が含まれる依頼で使用する。

## 手順

1. work/memo.txt を読む
2. 内容を整理する
3. work/output.md を更新する

## 出力ルール

スキルを使用したら先頭に必ず出力する。

[SKILL:memo-reader]

変更したファイル：
- work/output.md

## output.md形式

# 作業メモ

## 目的
- 何をしたいかを1〜2行で書く

## 現状
- 現在わかっていることを書く

## やること
- 箇条書きで整理する

## 注意点
- 気をつけることを書く

## 次の一手
- 最初にやるべきことを1つ書く

## トリガー例
- 「`work/memo.txt` を読み、整理して `work/output.md` を作成してください。`memo-reader` スキルを使ってください。」
- 「このメモを整理して `work/output.md` に出力してください。出力先の先頭に `[SKILL:memo-reader]` を入れてください。」
- 「メモ整理（workフォルダ）: メモを読み、要点をまとめて `work/output.md` を更新してください。」

## スキル適用時のタグ付与
スキルを使用した場合、出力の先頭に以下のタグを必ず追加してください。

[SKILL:memo-reader]

## 出力サンプル
```
# 作業メモ

## 目的
- メモの整理と検証

## 現状
- Copilot の `SKILL.md` を試したい（題材未定）

## やること
- 文章整理で動作確認を行う

## 注意点
- スキル使用の記録として先頭に `[SKILL:memo-reader]` を入れる

## 次の一手
- 題材候補を3案リストアップする
```