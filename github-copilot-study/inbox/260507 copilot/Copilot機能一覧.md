はい、最近の [GitHub Copilot](https://github.com/features/copilot?utm_source=chatgpt.com) はかなり「AI開発プラットフォーム化」していて、
単なる補完以外に、

* instructions
* agents
* skills
* prompts
* hooks
* MCP
* coding agent

など、多層的な仕組みがあります。

かなり整理が難しいので、実務での役割ベースでまとめます。

---

# 現在のCopilot系の主要機能群

大きく分けると：

| 分類            | 役割            |
| ------------- | ------------- |
| 基本AI機能        | 補完・Chat       |
| Instructions系 | AIへの恒久ルール     |
| Agent系        | AI人格・役割       |
| Skills系       | 特定タスク知識       |
| Prompt系       | 再利用タスク        |
| Tool連携        | CLI/MCP/Hooks |
| 自動化系          | Coding Agent  |

です。

---

# 1. copilot-instructions.md

最重要の基本機能です。

## 役割

「このプロジェクトではどう振る舞うべきか」をAIへ教える。 ([GitHub Docs][1])

## 配置

```text
.github/copilot-instructions.md
```

## 用途

* コーディング規約
* アーキテクチャ
* 禁止事項
* 命名規則
* 使用ライブラリ
* テスト方針

## 例

```md
- C# 7.3 を使用
- async/awaitは禁止
- Logger必須
- WinFormsを使用
- Newtonsoft.Jsonは禁止
```

---

# 2. *.instructions.md

より細かい条件別ルール。

## 例

```text
frontend.instructions.md
python.instructions.md
```

## 役割

特定ファイル群にだけ適用。

## 例

```yaml
applyTo: "**/*.py"
```

Python時だけ適用。

---

# 3. AGENTS.md

最近かなり重要。

## 役割

「AIエージェントの行動方針」。 ([SIOS Tech Lab][2])

## 配置

```text
AGENTS.md
```

または：

```text
.github/agents/
```

## 特徴

instructions が：

* 「何を守るか」

なのに対して、

AGENTS は：

* 「どう作業するか」

です。

---

## AGENTS.md に書くもの

### 作業手順

```md
1. テスト実行
2. lint確認
3. 修正
4. 再テスト
```

### 禁止事項

```md
- DB migration禁止
- 本番API禁止
```

### 実行コマンド

```md
npm test
dotnet build
```

---

# 4. Custom Agents（*.agent.md）

かなり新しい。

## 役割

「特定専門AIを作る」。 ([Zenn][3])

## 例

* SecurityReviewer.agent.md
* TestWriter.agent.md
* Refactor.agent.md

---

## イメージ

### 通常Copilot

汎用AI

### Custom Agent

専門職AI

---

## 例

```md
あなたはC# WinForms専門レビュアーです。
Dispose漏れを重点確認してください。
```

---

# 5. SKILL.md（Agent Skills）

最近かなり注目。

## 役割

「特定タスク知識をモジュール化」。 ([GitHub Docs][4])

---

## 配置

```text
.github/skills/
    testing/SKILL.md
    architecture/SKILL.md
```

---

## 最大の特徴

必要時だけ読み込まれる。

つまり：

* instructions
  → 常時読む

* skills
  → 必要時だけ読む

---

# SKILL.md の実務的意味

これがかなり重要です。

例えば：

## testing/SKILL.md

```md
Vitest + React Testing Library のテスト方針
```

↓

テスト関連タスク時だけ読む。

---

# skills の利点

巨大 instructions の問題を解決。

---

## 悪い例

巨大：

```md
copilot-instructions.md
```

↓

AIが忘れる
ノイズ化。

---

## 良い例

分割：

```text
skills/
    testing/
    db/
    ui/
    logging/
```

---

# 6. Prompt Files（*.prompt.md）

かなり便利。

## 役割

「再利用可能プロンプト」。

([Zenn][3])

---

## 例

```text
create-api.prompt.md
review.prompt.md
```

---

## 用途

毎回同じ依頼を定型化。

---

## 例

```md
# API生成

以下を実施：
- DTO生成
- Interface生成
- Logger追加
- Test生成
```

---

# 7. Hooks

かなりエージェント寄り。

## 役割

AI処理前後にスクリプト実行。

---

## 例

### Before Hook

* 環境確認
* lint

### After Hook

* test実行
* format

---

# 8. MCP（Model Context Protocol）

最近の超重要ワード。

---

## 役割

AIへ外部ツールを接続。

---

## 例

Copilotが：

* DB参照
* GitHub参照
* Slack参照
* Docs検索
* Figma参照

など可能。

---

# 実質的には

「AI用プラグイン規格」。

---

# 9. Copilot Coding Agent

最近の本命。

---

## 役割

AIが：

* Issue読む
* 実装
* 修正
* test
* commit
* PR作成

までやる。

---

# 従来との違い

## 従来

人間主導：

```text
人 → AI補助
```

## Agent

AI主導：

```text
人 → 指示
AI → 実装
```

---

# 10. Copilot Workspace

巨大機能。

---

## 役割

Issueから：

* 設計
* 実装計画
* 修正
* PR

を生成。

---

# 11. Copilot CLI

ターミナルAI。

## 例

```bash
gh copilot suggest
```

---

# 12. Spaces

最近追加。

## 役割

「共有知識ベース」。

([GitHub][5])

---

## 例

チーム共通：

* 設計書
* API仕様
* 運用ルール

をAI共有。

---

# 13. awesome-copilot

コミュニティ集。

## 内容

* agents
* skills
* hooks
* workflows

など。

([GitHub][6])

---

# 実務での整理（超重要）

現在のCopilotは：

| レイヤ          | 役割     |
| ------------ | ------ |
| instructions | 全体ルール  |
| skills       | 専門知識   |
| agents       | AI人格   |
| prompts      | 作業テンプレ |
| hooks        | 自動処理   |
| MCP          | 外部接続   |
| coding agent | 自律実装   |

という構造。

---

# ユーザーさんの用途だと特に相性良いもの

かなり相性が良いです。

---

## ① copilot-instructions.md

例えば：

```md
- .NET Framework 4.7.2
- WinForms
- Logger必須
- async禁止
- C#7.3対応
```

これだけでも精度かなり上がります。

---

## ② skills

かなりおすすめ。

### 例

```text
skills/
    oracle-db/
    ffmpeg/
    logging/
    winforms/
    installshield/
```

---

## ③ custom agent

### 例

```text
VBNetReviewer.agent.md
```

* Nullチェック
* Dispose
* OracleConnection Close
* Logger

を重点確認。

---

## ④ prompt

### 例

```text
create-detail-design.prompt.md
```

↓

詳細設計書を毎回同じ形式で生成。

---

# Claude Code系との関係

最近かなり重要。

実は：

* AGENTS.md
* SKILL.md

は、Claude Code互換を意識した設計になっています。 ([SIOS Tech Lab][7])

つまり：

```text
GitHub Copilot
Claude Code
```

両対応を狙いやすい。

---

# 今後の流れ

今の流れは：

## 2023

補完AI

↓

## 2024

Chat AI

↓

## 2025

Agent AI

↓

## 2026

自律開発環境

になっています。

---

# 個人的に実務で重要と思う順

現時点だと：

| 重要度 | 機能                      |
| --- | ----------------------- |
| S   | copilot-instructions.md |
| S   | Agent Mode              |
| A   | SKILL.md                |
| A   | AGENTS.md               |
| A   | Prompt Files            |
| B   | MCP                     |
| B   | Hooks                   |
| B   | Custom Agents           |

くらいです。

---

# 参考

* [GitHub Copilot Docs](https://docs.github.com/en/copilot?utm_source=chatgpt.com)
* [Agent Skills Docs](https://docs.github.com/en/copilot/concepts/agents/about-agent-skills?utm_source=chatgpt.com)
* [Custom Instructions Docs](https://docs.github.com/copilot/customizing-copilot/adding-custom-instructions-for-github-copilot?utm_source=chatgpt.com)
* [Awesome Copilot](https://github.com/github/awesome-copilot?utm_source=chatgpt.com)

[1]: https://docs.github.com/copilot/customizing-copilot/adding-custom-instructions-for-github-copilot?utm_source=chatgpt.com "Adding repository custom instructions for GitHub Copilot"
[2]: https://tech-lab.sios.jp/archives/51144?utm_source=chatgpt.com "copilot-instructions.md と AGENTS.md、どっちに何を書く ..."
[3]: https://zenn.dev/thinkingsinc/articles/cd59e1ffc4f842?utm_source=chatgpt.com "GitHub Copilotにおける Agent / Instructions / Prompt の整理 ..."
[4]: https://docs.github.com/en/copilot/concepts/agents/about-agent-skills?utm_source=chatgpt.com "About agent skills - GitHub Docs"
[5]: https://github.com/github/awesome-copilot/blob/main/docs/README.skills.md?utm_source=chatgpt.com "awesome-copilot/docs/README.skills.md at main"
[6]: https://github.com/github/awesome-copilot?utm_source=chatgpt.com "github/awesome-copilot"
[7]: https://tech-lab.sios.jp/archives/51023?utm_source=chatgpt.com "GitHub Copilot設定5種を網羅！生産性を最大化する使い分け術"
