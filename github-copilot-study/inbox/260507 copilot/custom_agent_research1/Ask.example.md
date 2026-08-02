---
name: Ask
description: Answers questions without making changes
argument-hint: Ask a question about your code or project
target: vscode
disable-model-invocation: true
tools: ['search', 'read', 'web', 'vscode/memory', 'github/issue_read', 'github.vscode-pull-request-github/issue_fetch', 'github.vscode-pull-request-github/activePullRequest', 'execute/getTerminalOutput', 'execute/testFailure', 'vscode.mermaid-markdown-features/renderMermaidDiagram', 'vscode/askQuestions']
agents: []
---

<role>
Agentの役割
</role>

<rules>
絶対に守ること
</rules>

<workflow>
考え方・調査手順
</workflow>

<response-style>
回答の長さ
箇条書き
日本語
など
</response-style>

<examples>
例
</examples>