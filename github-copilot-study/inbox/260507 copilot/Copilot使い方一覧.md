[GitHub Copilot公式サイト](https://github.com/features/copilot?utm_source=chatgpt.com) は、単なる「コード補完」だけではなく、かなり幅広い用途があります。
現在の主な使い方・機能を、実務寄りに整理すると以下のようになります。

---

# GitHub Copilot の主な機能一覧

## 1. コード補完（Inline Completion）

最も基本的な機能です。

### できること

* 次の1行を予測
* 複数行まとめて生成
* 関数全体を生成
* コメントからコード生成

### 例

```python
# 指定フォルダのmp4ファイル一覧を取得
```

↓

```python
files = list(Path(target_dir).glob("*.mp4"))
```

### 特徴

* Tabで採用
* 入力途中でも候補表示
* 文脈（周辺コード）を見て補完

---

# 2. Copilot Chat（AIチャット）

VSCode / Visual Studio 内でAIと会話できます。

## できること

### コード説明

「このコード何してる？」

### バグ解析

「NullReferenceException の原因は？」

### リファクタ提案

「クラス分割したい」

### 設計相談

「WinFormsで設定管理どうする？」

### テストコード生成

「xUnitテスト作って」

### SQL生成

「このテーブル構成でJOIN文作って」

### 正規表現生成

「yyyymmddを抜き出すregex」

---

# 3. コメント → コード生成

かなり実用的です。

## 例

```csharp
// iniファイルを読み込んで Dictionary<string,string> に変換
```

↓

自動生成。

---

# 4. 関数単位の生成

## 例

```python
def get_movie_duration(path:str) -> float:
```

↓

中身を生成。

---

# 5. テストコード生成

## 対応例

* NUnit
* xUnit
* MSTest
* pytest
* unittest

## 例

「このクラスの異常系テスト作成」

---

# 6. リファクタリング支援

## 例

* 長いメソッド分割
* 名前改善
* 重複コード削除
* async化
* LINQ化
* nullable対応

---

# 7. ドキュメント生成

## 例

### XMLコメント

```csharp
/// <summary>
```

↓

自動生成。

### Markdown設計書

* README
* API仕様
* 要求仕様
* 詳細設計

など。

---

# 8. エラー解析支援

## 例

エラー全文を貼ると：

* 原因推定
* 修正候補
* 関連箇所
* 再現条件

などを提案。

---

# 9. Git操作支援

## 例

* Commit Message生成
* PR説明生成
* 差分要約
* コードレビュー

---

# 10. コードレビュー支援（Copilot Code Review）

Pull Requestに対して：

* バグ候補
* セキュリティ問題
* 可読性問題
* nullチェック不足

などを指摘。

---

# 11. CLI支援（Copilot in CLI）

ターミナル上でAI利用。

## 例

```bash
gh copilot suggest "zipファイルを再帰展開"
```

---

# 12. 正規表現生成

かなり便利です。

## 例

「hh:mm:ss を検出」

↓

regex生成。

---

# 13. SQL生成

## 例

* CREATE TABLE
* JOIN
* UPDATE
* UPSERT
* INDEX提案

---

# 14. API呼び出しコード生成

## 例

* REST API
* HttpClient
* requests
* fetch
* axios

---

# 15. UIコード生成

## 対応例

* WinForms
* WPF
* HTML
* React
* Flutter

---

# 16. 学習用途

かなり強いです。

## 使い方

* 「このコードを初心者向けに説明」
* 「VB.NET版に変換」
* 「C#7.3対応にして」
* 「async使わない版」

など。

---

# 17. コード変換

## 例

* VB.NET → C#
* Python → C#
* Java → Python

---

# 18. 設計相談

実はかなり使われています。

## 例

* クラス構成
* namespace設計
* ロガー設計
* 非同期設計
* DI導入
* DB構成

---

# 19. ファイル全体理解

最近かなり強化されています。

## 例

* 「このプロジェクト構成説明して」
* 「依存関係教えて」
* 「未使用クラスある？」

---

# 20. Workspace / Agent系（新機能系）

最近は「AIエージェント化」が進んでいます。

## できること

* issueから実装
* 複数ファイル修正
* テスト実行
* 修正提案
* PR生成

かなり「半自動開発」に近づいています。

---

# Visual Studio / VSCodeでの実務利用パターン

ユーザーさんのような：

* C#
* VB.NET
* Python
* WinForms
* ffmpeg
* 設計書作成
* ログ設計
* クラス分割

の用途だと、特に相性が良いのは：

* コメント駆動実装
* リファクタ
* 例外処理追加
* Logger埋め込み
* 設計書md生成
* Regex生成
* INI/JSON処理
* WinFormsイベント生成

です。

---

# 実務で便利な使い方（おすすめ）

## ① まずコメントを書く

```csharp
// Drag&Dropされたファイルを検証してTextBoxに設定
```

→ Copilot補完。

---

## ② クラス構成だけ先に書く

```csharp
class AppSettingsManager
{
}
```

→ メソッド候補が出る。

---

## ③ エラー全文をChatへ貼る

かなり有効です。

---

## ④ 「このコードにLogger追加」

も便利です。

---

# 注意点

## 間違うことは普通にある

特に：

* async
* IDisposable
* DB
* セキュリティ
* マルチスレッド

は要レビュー。

---

# 特に強い言語

かなり精度が高い：

* Python
* TypeScript
* JavaScript
* C#
* Java

そこそこ：

* VB.NET
* VBA
* PowerShell

---

# 料金系

現在は主に：

* Free
* Pro
* Business
* Enterprise

があります。

個人開発なら通常は Pro が主流です。

---

# 関連サービス

* [GitHub Copilot Docs](https://docs.github.com/copilot?utm_source=chatgpt.com)
* [Visual Studio Copilot Docs](https://learn.microsoft.com/visualstudio/ide/visual-studio-github-copilot-extension?utm_source=chatgpt.com)
* [VSCode Copilot Docs](https://code.visualstudio.com/docs/copilot/overview?utm_source=chatgpt.com)
