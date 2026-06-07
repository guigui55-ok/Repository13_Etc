---
name: github-actions-failure-debugging
description: Guide for debugging failing GitHub Actions workflows. Use this when asked to debug failing GitHub Actions workflows.
---

To debug failing GitHub Actions workflows in a pull request, follow this process, using tools provided from the GitHub MCP Server:

1. Use the `list_workflow_runs` tool to look up recent workflow runs for the pull request and their status
2. Use the `summarize_job_log_failures` tool to get an AI summary of the logs for failed jobs, to understand what went wrong without filling your context windows with thousands of lines of logs
3. If you still need more information, use the `get_job_logs` or `get_workflow_run_logs` tool to get the full, detailed failure logs
4. Try to reproduce the failure yourself in your own environment.
5. Fix the failing build. If you were able to reproduce the failure yourself, make sure it is fixed before committing your changes.


========================

---
name: github-actions-failure-debugging
description: 失敗している GitHub Actions ワークフローをデバッグするためのガイド。GitHub Actions のワークフロー失敗を調査するよう依頼された場合に使用します。
---

プルリクエスト内で失敗している GitHub Actions ワークフローをデバッグする場合は、GitHub MCP Server が提供するツールを使い、次の手順で進めます。

1. `list_workflow_runs` ツールを使用して、対象プルリクエストに関連する最近のワークフロー実行履歴とそのステータスを確認する。

2. `summarize_job_log_failures` ツールを使用して、失敗したジョブのログを AI に要約させる。これにより、何千行ものログでコンテキストを埋めることなく、問題の概要を把握する。

3. さらに詳細な情報が必要な場合は、`get_job_logs` または `get_workflow_run_logs` ツールを使用して、完全な失敗ログを取得する。

4. 自分のローカル環境（または利用可能な実行環境）で、失敗を再現できるか試す。

5. ビルド失敗を修正する。もし失敗を再現できた場合は、変更をコミットする前に、確実に問題が解消されていることを確認する。

========================
# 構成
```
name: skill-name
description : 「このスキルをいつ使うか（発火条件）」を簡潔に書く
---

目的・基本方針
どういう思想・順番で実施するか（実行プロセス）

実行時のルール・前提条件

手順は番号付き推奨（ただし必須ではない）

手順
1.
2.
3.

補足（必要なら）
```

特に description は重要で、**人向け説明というより Agent に「このスキルを選ばせる条件を書く場所」**という理解にすると、スキルが増えても管理しやすくなります。

=========================
なので以前ユーザーさんが言っていた、

スキルって依頼内容によって勝手に使用される想定だった

という感覚は方向性として合っています。

ただ実際には現状（2026時点）の Copilot は、

description が弱い
Agent 側の指示が弱い
スキル数が多い
要求が曖昧

だと選んでくれないことがあります。

だから実運用では、

AGENT.md
必要に応じて skills を選択してください。

や、

この作業では memo-organizer skill を使用してください

みたいな補助を書くことがまだ結構あります。

