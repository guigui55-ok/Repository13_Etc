User:
Copilot の MCP について、常駐企業先（職場）で使用できないかと考えています。
ただ、職場は非常にセキュリティが厳しく、外部のソフトウェアなど請負元企業にすべて許可が必要です。
なので、最小限の構成で使用したいと考えています。
このような環境の時、Copilot MCP はどのような用途で、そして、どのような構成で使用するのが効率がよいでしょうか？

案件内容としては、既存大規模ソフトウェア（C#クライアント＋サーバー×2[SqlServer、C#Webサービス]）の機能改善・追加が継続的にある案件で、仕様書がたくさんあるプロジェクトです。
案件対応の例としては、クライアントやサーバーの機能の改善と追加ですが、その際に、仕様書との整合性を綿密に調べたり、既存ソースとの整合性や影響をよく調べたうえ、詳細設計書を記載して、実装・テストをします。

Chat:
結論としては、**いきなりMCPを本格導入するより、「Copilot標準機能＋必要最小限のローカルMCP」から始める**のがよいです。

GitHub公式では、MCPはCopilotに外部ツール・外部データソースを接続する仕組みです。VS CodeではMCPサーバーを追加して、ファイル操作、DB、外部APIなどをCopilotから使えるようにできます。([GitHub Docs][1])

## おすすめ構成

### 第1段階：MCPなし、Copilot標準機能のみ

職場の許可を取りやすい最小構成です。

```text
VS Code / Visual Studio
  └ GitHub Copilot Chat
       └ 開いているソース
       └ 選択範囲
       └ ワークスペース内検索
       └ 手動で貼り付けた仕様書抜粋
```

用途は以下に絞るのが安全です。

* 既存C#コードの処理説明
* 影響範囲調査の観点出し
* 詳細設計書のたたき台作成
* テスト観点の洗い出し
* SQL Serverテーブル・API・画面処理の関係整理
* 仕様書の抜粋を貼って、ソースとの差分確認

この段階でもかなり使えます。
厳しい職場なら、まずここまでで申請するのが現実的です。

## 第2段階：ローカル・読み取り専用MCP

MCPを使うなら、最初はこれが一番おすすめです。

```text
VS Code / Visual Studio
  └ GitHub Copilot Chat
       └ ローカルMCPサーバー
            ├ 仕様書検索
            ├ ソース検索
            ├ 関連ファイル一覧取得
            └ grep / ripgrep 相当の検索
```

ポイントは、**外部API連携なし、DB接続なし、書き込みなし**です。

許可申請しやすい説明としては、

> MCPは外部サービス連携ではなく、社内PC上の許可されたプロジェクトフォルダを読み取り専用で検索するために使用する。
> ソース・仕様書・設計書の検索補助を目的とし、ファイル更新、DB更新、外部送信、外部API呼び出しは行わない。

という形がよいです。

VS CodeのMCP設定はユーザーレベルまたはワークスペースレベルで構成でき、リポジトリ単位では `.vscode/mcp.json` に設定できます。([GitHub Docs][2])

## この案件でMCPが効く用途

一番効果が出るのは、**仕様書と既存ソースの突き合わせ**です。

たとえばMCPに以下のような読み取り専用ツールを持たせます。

```text
search_specs(keyword)
search_source(keyword)
find_related_files(functionName / className / tableName)
get_file_excerpt(path, lineRange)
search_sql_objects(tableName / procedureName)
```

すると、Copilotにこう聞けます。

```text
「受注登録機能の仕様変更について、
仕様書上の関連箇所と、既存C#クライアント、Webサービス、SQL Server側の
影響候補を一覧化してください」
```

または、

```text
「この詳細設計書案について、
既存仕様書・既存ソースと矛盾しそうな点を洗い出してください」
```

この用途なら、MCPは「勝手に実装する道具」ではなく、**調査補助・検索補助**として使えます。

## 避けた方がよい構成

厳しい職場では、最初から以下は避けた方がよいです。

* 外部SaaS連携MCP
* GitHubリポジトリ操作MCP
* DB更新可能なMCP
* ファイル書き込み可能なMCP
* 野良のMCPサーバー
* npm / pip で多数の依存パッケージを入れる構成
* 自動でPR作成、Issue更新、コミットする構成

GitHub MCP Serverは、GitHub上のリポジトリ、Issue、PRなどを操作できるため便利ですが、職場では権限・監査・許可範囲の説明が必要になります。([GitHub][3])

## 最小構成の案

個人的には、この順番がよいです。

```text
案1：Copilot標準機能のみ
↓
案2：ローカル読み取り専用MCP
↓
案3：GitHub MCPは必要になってから検討
↓
案4：DB接続MCPや書き込み系MCPは原則使わない
```

## 申請時の説明に使える文面

```text
GitHub Copilot MCPの利用目的は、既存ソースコードおよび仕様書の調査効率化です。

初期構成では、ローカルPC上の許可されたプロジェクトフォルダに対して、
読み取り専用の検索機能のみを提供します。

MCP経由で以下は行いません。
・ファイルの作成、更新、削除
・DBへの接続、更新
・外部Web APIへの接続
・GitHub Issue / PR / Repository の更新
・社外サービスへの任意データ送信

用途は、仕様書と既存ソースの整合性確認、影響範囲調査、
詳細設計書作成時の調査補助、テスト観点抽出に限定します。
```

## 実務上のおすすめ

この案件では、MCPは「実装を任せる」よりも、**調査漏れを減らす道具**として使うのが一番安全で効果的です。

特に有効なのは、

```text
仕様変更
→ 関連仕様書検索
→ 関連画面・クラス・Webサービス・SQL検索
→ 影響範囲一覧
→ 詳細設計書の章立て作成
→ テスト観点作成
```

この流れです。

まずは **Copilot標準機能のみで申請**し、次に **ローカル読み取り専用MCP** を追加申請する形が、セキュリティが厳しい職場では一番通しやすいと思います。

[1]: https://docs.github.com/en/copilot/concepts/context/mcp?utm_source=chatgpt.com "About Model Context Protocol (MCP) - GitHub Copilot"
[2]: https://docs.github.com/copilot/customizing-copilot/using-model-context-protocol/extending-copilot-chat-with-mcp?utm_source=chatgpt.com "Extending GitHub Copilot Chat with Model Context ..."
[3]: https://github.com/github/github-mcp-server?utm_source=chatgpt.com "GitHub's official MCP Server"
