---
name: csharp-version-check-investigation
description: C#プロジェクト内のサーバーバージョン判定処理の使用箇所を調査し列挙する
---

# C# バージョン判定処理調査

## 目的
サーバーとのバージョン判定処理が、どこで使われているかを調査する。

## 調査キーワード
- Version
- ServerVersion
- CheckVersion
- CompareVersion
- GetVersion
- IsSupported
- Compatible
- UpdateRequired
- バージョン
- 版数
- サーバー
- 互換

## 手順
1. バージョン関連キーワードを検索する
2. 判定メソッド・プロパティを特定する
3. 呼び出し元を列挙する
4. 判定結果による分岐を確認する
5. 仕様書に該当記述があるか確認する
6. 結果を一覧表で出力する

## 出力形式
| No | ファイル | クラス/メソッド | 判定内容 | 使用箇所 | 条件分岐 | 仕様書対応 | 備考 |