---
name: csharp-file-write-investigation
description: C#プロジェクト内のファイル書き込み処理を調査し、暗号化追加時の影響箇所を列挙する
---

# C# ファイル書き込み処理調査

## 目的
ファイル出力処理に暗号化を追加する前提で、既存の書き込み処理を調査し、変更候補を列挙する。

## 調査対象
以下を優先して検索する。

- File.WriteAllText
- File.WriteAllBytes
- File.AppendAllText
- StreamWriter
- FileStream
- BinaryWriter
- XmlWriter
- JsonSerializer
- CSV 出力処理
- Save / Export / Output / Write / CreateFile などの命名

## 手順
1. 書き込みAPIを検索する
2. 書き込み先ファイルの種類を確認する
3. 呼び出し元を追跡する
4. 仕様書に対応する記述があるか確認する
5. 暗号化を追加すべき箇所と、追加すべきでない箇所を分ける
6. 結果を一覧表で出力する

## 出力形式
| No | ファイル | クラス/メソッド | 書き込み内容 | 呼び出し元 | 暗号化対象候補 | 備考 |