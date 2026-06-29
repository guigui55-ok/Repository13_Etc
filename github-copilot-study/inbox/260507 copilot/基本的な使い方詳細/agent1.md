GitHub Copilot の「Custom Agents」は、GitHub Copilot に対して
「この役割・方針・使えるツール・動作ルールで振る舞ってほしい」
を定義できる仕組みです。

簡単に言うと、

* 「設計専用AI」
* 「セキュリティレビュー専用AI」
* 「WinForms保守専用AI」
* 「Pythonテストコード生成専用AI」

のような “専門人格” を作れる機能です。 ([GitHub Docs][1])

---

# まず整理：Custom Agents は何を解決する？

通常のCopilot Chatだと、

* 毎回前提を説明する
* 毎回「C# 7.3で」「WinFormsで」「既存構造維持で」
  などを書く
* 毎回「勝手にasync化しないで」と指示する

必要があります。

Custom Agent を使うと、その前提を固定化できます。

つまり：

```text
このAgentは、
- C# WinForms専用
- .NET Framework 4.7.2前提
- async禁止
- AppLogger使用
- 既存構造を維持
- 詳細ログ必須
```

のようなルールを持ったCopilotを作れます。

かなり、あなたの開発スタイルと相性が良いです。

---

# Custom Agents のイメージ

例えば：

| Agent名                     | 用途                 |
| -------------------------- | ------------------ |
| `planner.agent.md`         | 要求仕様・設計書生成         |
| `legacy-winforms.agent.md` | 既存WinForms保守       |
| `vbnet-db.agent.md`        | Oracle/VB.NET DB処理 |
| `test-reviewer.agent.md`   | テスト観点レビュー          |
| `python-ffmpeg.agent.md`   | ffmpegツール専用        |

のように分けます。

---

# 何が設定できる？

Custom Agentでは主に：

* 振る舞い
* 使用ツール
* MCP
* モデル
* 指示
* handoff（別Agentへ引き継ぎ）

を設定できます。 ([Visual Studio Code][2])

---

# 実体は何？

基本は `.agent.md` ファイルです。

例えば：

```text
.github/agents/winforms.agent.md
```

のようなファイルを作ります。 ([GitHub Docs][3])

---

# 最小構成例

例えばあなた向けなら：

```md
---
name: Legacy WinForms Engineer
description: C# WinForms maintenance and extension specialist
tools:
  - codebase
  - terminal
  - editFiles
---

# Role

You are a senior C# WinForms engineer.

# Rules

- Target framework: .NET Framework 4.7.2
- Keep existing structure
- Avoid async unless requested
- Use AppLogger
- Prefer explicit classes over lambda expressions
- Preserve backward compatibility

# Coding style

- Add detailed logging
- Avoid unnecessary refactoring
- Prefer readable code
```

こんな感じです。

---

# VSCodeでどう使う？

VSCode の Copilot Chat の Agent選択から切り替えます。 ([Visual Studio Code][2])

例えば：

```text
Ask | Edit | Agent
```

の「Agent」モードで、

```text
Legacy WinForms Engineer
```

を選ぶイメージです。

---

# Agent と Prompt.md の違い

ここかなり重要です。

| 機能                  | 目的          |
| ------------------- | ----------- |
| prompt.md           | 一時的な指示      |
| custom instructions | リポジトリ共通ルール  |
| skills.md           | 再利用機能       |
| Custom Agent        | “人格・役割” の切替 |

つまり：

* prompt.md → 単発
* skills → 能力部品
* agent → 専門職AI

です。

---

# あなたの場合、かなり向いている用途

あなたの開発スタイルを見ると：

* 既存保守
* レガシーWinForms
* 詳細設計重視
* ログ重視
* async抑制
* 明示クラス構成
* Oracle/VB.NET/C#

など、独自ルールがかなり多いので、
Custom Agent の効果が大きいです。

特に：

## ① 保守用Agent

```text
既存コードを壊さない
```

を強制できる。

---

## ② 設計書Agent

```text
要求仕様書
基本設計書
詳細設計書
```

生成専用Agent。

---

## ③ テストAgent

```text
試験観点
異常系
ログ確認
```

重視のAgent。

---

# MCPとの関係

かなり重要です。

Custom Agent は MCP と組み合わせできます。 ([GitHub Docs][3])

つまり：

```text
このAgentは
- Oracle DB MCP
- Git MCP
- FileSystem MCP
だけ使用可能
```

みたいな制御ができます。

これが「業務特化AI」になります。

---

# Agent Mode との違い

これも混同しやすいです。

| 機能           | 意味      |
| ------------ | ------- |
| Agent Mode   | 自律実行モード |
| Custom Agent | AI人格定義  |

です。

つまり：

* Agent Mode = 行動方式
* Custom Agent = 性格・専門性

です。 ([GitHub Docs][4])

---

# 実際のおすすめ構成（あなた向け）

あなたなら最初は：

```text
.github/
 ├ agents/
 │   ├ winforms.agent.md
 │   ├ planner.agent.md
 │   ├ vbnet.agent.md
 │   └ python-tool.agent.md
```

くらいがかなり良いです。

---

# 特におすすめなAgent

あなた向けに特に有効なのは：

## WinForms保守Agent

```text
- .NET Framework 4.7.2
- WinForms
- AppLogger mandatory
- async禁止
- 既存構造維持
```

---

## 設計書Agent

```text
- Markdown出力
- 要求仕様→基本設計→詳細設計
- テーブル形式
- 機能ID付与
```

---

## ffmpeg/Python Agent

```text
- pathlib使用
- logger mandatory
- ffmpeg-python使用
- 例外ログ必須
```

---

# 注意点

まだ比較的新しい機能で、
環境によって：

* 使用可能範囲
* UI
* organization agent
* cloud agent連携

が変わります。 ([GitHub Docs][3])

特に：

* VSCode
* GitHub.com
* Visual Studio
* Enterprise

で少し仕様差があります。

---

# まず最初にやると良いこと

おすすめ順：

1. custom instructions
2. prompt.md
3. skills
4. Custom Agent
5. MCP連携

です。

あなたは既に：

* prompt設計
* skills
* Agent Mode
* MCP

に興味を持っているので、
次の段階としてかなり自然です。

[1]: https://docs.github.com/en/copilot/concepts/agents/cloud-agent/about-custom-agents?utm_source=chatgpt.com "About custom agents"
[2]: https://code.visualstudio.com/docs/copilot/customization/custom-agents?utm_source=chatgpt.com "Custom agents in VS Code"
[3]: https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/create-custom-agents?utm_source=chatgpt.com "Creating custom agents for Copilot cloud agent"
[4]: https://docs.github.com/copilot/concepts/agents/coding-agent/about-coding-agent?utm_source=chatgpt.com "About GitHub Copilot cloud agent"
