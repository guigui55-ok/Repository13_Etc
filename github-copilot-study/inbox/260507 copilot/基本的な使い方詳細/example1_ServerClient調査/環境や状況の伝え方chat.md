
---

## ① プロジェクト情報（技術構成・制約） → agent / instructions

これは毎回効く情報です。

例：

* C#.NET Framework 4.7
* WinForms
* クライアントサーバー構成
* サーバー：ASP.NET SOAP(asmx)
* Oracle DB
* Visual Studio 2019
* 非同期使用制限あり
* ライブラリ追加禁止
* .NET Core移行予定なし

こういうもの。

理由：
これらは調査方法・提案方法に常時影響するからです。

例：

❌ 毎回 Prompt に書く

```text
この案件は.NET Framework4.7です…
SOAP(asmx)です…
```

↓

毎回長い。

---

推奨：

```text
agents/
└─ enterprise-csharp-agent/
    ├─ AGENT.md
    └─ project_context.md
```

AGENT.md

```md
# Enterprise C# Investigation Agent

前提:
- Client/Server
- C#.NET Framework 4.7
- Server: ASP.NET SOAP(asmx)
- OracleDB
- Visual Studio 2019

制約:
- ライブラリ追加禁止
- .NET Core置換禁止
- SQL変更時は影響範囲確認

出力:
- ファイル
- クラス
- メソッド
- 呼び出し元
```

これなら全部の作業に効きます。

---

## ② 作業背景・目的 → prompt

ここが今回の質問の本命です。

> 顧客先でプログラム使用時に情報漏洩についてのセキュリティを強化するため

これは**書いた方が良いです。かなり重要です。**

理由：
同じ改修でも調査観点が変わるからです。

例。

依頼①

```text
ファイル出力処理へ暗号化追加
目的：情報漏洩対策
```

調査対象：

* 一時ファイル
* CSV
* ログ
* Temp
* キャッシュ
* 平文保存

---

依頼②

```text
ファイル出力処理へ暗号化追加
目的：通信速度改善
```

調査対象：

* サイズ
* 圧縮
* CPU負荷

全然違います。

だから背景は Prompt に入れる。

例：

```text
目的:
顧客先利用時の情報漏洩対策

変更内容:
ファイル出力へ暗号化追加

調査:
- 平文出力箇所
- 一時保存
- ログ出力
- 復号タイミング

成果物:
investigation.md
```

---

## ③ 調査観点・チェック項目 → skill

これは使い回したい手順。

例：

```text
skills/
├─ file-encryption-investigation
├─ security-impact-analysis
├─ soap-interface-analysis
├─ oracle-update-impact
```

skill：

```md
セキュリティ調査時は確認:

□ Temp保存
□ ログ
□ SOAP送信
□ XML保存
□ Exception出力
□ 設定ファイル
```

---

なので、今回の例なら最終的にはこう分けます。

```text
repo/
├─ .github/
│   ├─ AGENT.md
│   │    ← 技術構成・恒久制約
│   │
│   ├─ skills/
│   │   ├─ file-write-investigation/
│   │   ├─ security-impact-analysis/
│   │   └─ soap-interface-investigation/
│   │
│   └─ prompts/
│       └─ encrypt-export.prompt.md
```

役割：

```text
AGENT
「この案件は何者か」

SKILL
「どう調査するか」

PROMPT
「今回はなぜやるか」
```

あなたの例の「情報漏洩対策」は、技術仕様ではなく**要求背景**なので Prompt 側に置くのが一番効きます。