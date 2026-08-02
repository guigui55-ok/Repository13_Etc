別エージェントの呼び出し（サブエージェントの自動呼び出し）
```yaml
tools: ['search', 'read', 'web', 'vscode/askQuestions', 'agent']
agents: ['Deep Ask']
```

handoffs
ハンドオフは回答後に切り替えボタンを表示し、ユーザーが実行を判断できる仕組みです。
```yaml
handoffs:
  - label: 詳細調査へ切り替え
    agent: Deep Ask
    prompt: この依頼について、必要な範囲で詳細な調査を行ってください。
    send: false
```
