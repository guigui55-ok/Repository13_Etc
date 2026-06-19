GitHub Copilot（VS Code）の `*.agent.md`（Custom Agent）の YAML ヘッダで設定する `tools` は、「この Agent が使えるツールを制限する設定」です。省略すると全ツール許可になります。([GitHub Docs][1])

まず結論です。

```yaml
---
tools: ["read", "edit", "search"]
---
```

のように記載します。

使える指定方法は次の3種類です。

```yaml
# 全許可（省略でも同じ）
tools: ["*"]

# 標準ツールだけ許可
tools: ["read", "search"]

# MCPツール含め個別指定
tools:
  - read
  - edit
  - search
  - my-mcp-server/tool-a
```

([GitHub Docs][1])

---

## 標準 Tool Alias 一覧（VS Codeでよく見るもの）

GitHub公式では「利用可能なツールは環境依存（VS Code・MCP・拡張機能追加で増える）」ですが、標準でよく使うものは次です。([GitHub Docs][1])

| tools値        | 用途             |
| ------------- | -------------- |
| `read`        | ファイル・コードを読む    |
| `edit`        | ファイル編集         |
| `search`      | ワークスペース検索      |
| `runCommands` | ターミナルコマンド実行    |
| `runTasks`    | VS Code Task実行 |
| `test`        | テスト実行          |
| `githubRepo`  | GitHubリポジトリ操作  |
| `changes`     | 差分確認           |
| `extensions`  | VS Code拡張利用    |
| `problems`    | Problemsパネル確認  |
| `usages`      | シンボル参照検索       |
| `fetch`       | 外部情報取得（環境依存）   |

※ 上の一部は利用環境やバージョン、Agent種別（ローカル/Cloud）で出たり出なかったりします。公式も「利用可能ツール一覧は環境依存」としています。([GitHub Docs][1])

---

## 一番確実な確認方法（おすすめ）

以前、`allowed-tools` の話もしていましたが、Agent の `tools` は VS Code 上で確認するのが一番確実です。

### 方法1（おすすめ）

1. `*.agent.md` を開く
2. `tools:` の箇所をクリック
3. **Configure Tools...** を押す
4. 利用可能ツール一覧がチェックボックスで出る

ここに表示されるものが、その環境で実際に指定可能な一覧です。([サーバーワークスエンジニアブログ][2])

---

## あなたの今の用途（メモ整理Agent）ならおすすめ

以前作っていたメモ整理系なら、まずはこれくらいで十分です。

```yaml
---
description: メモを整理して優先順位付けする
tools:
  - read
  - edit
  - search
---
```

調査専用 Agent なら編集禁止にして、

```yaml
tools:
  - read
  - search
```

にすると誤編集防止になります。

補足ですが、以前話していた `tools: [read, edit, search]` は「かなり定番構成」です。

[1]: https://docs.github.com/en/copilot/reference/custom-agents-configuration?utm_source=chatgpt.com "Custom agents configuration"
[2]: https://blog.serverworks.co.jp/getting-started-custom-agents-github-copilot-agent-mode?utm_source=chatgpt.com "はじめてのカスタムエージェント【GitHub Copilot Agent Mode編】"
