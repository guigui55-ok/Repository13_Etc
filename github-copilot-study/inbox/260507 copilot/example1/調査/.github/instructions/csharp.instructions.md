---
applyTo: "**/*.cs"
---

# C# 実装指示

## 対象環境

- .NET Framework 4.7 / 4.7.2
- WinForms
- 業務システム
- レガシーコードを含む

---

# 実装方針

- 既存実装に合わせる
- 可読性を優先する
- 新規ライブラリ追加は最小限
- 大規模変更は禁止
- 影響範囲を最小化する

---

# 命名

- 既存命名規則を優先
- public API変更は極力避ける

---

# WinForms

- Designer.cs を直接編集しない
- UIスレッドをブロックしない
- Invoke / BeginInvoke に注意する

---

# ファイル処理

- 文字コードに注意する
- Shift-JIS を使用している可能性がある
- BOM有無を維持する

---

# 非同期処理

- 不要な async 化は禁止
- 既存同期処理を優先する

---

# ログ出力

以下形式を使用する。

```csharp
_logger.Info("message");
_logger.Error(ex, "message");