User:
次はskills.mdについて知りたいです。 今回のような例では有効ですか？

AI:
はい。
そして結論から言うと、

```text id="9g2yy1"
ユーザーさんの環境は Skill とかなり相性が良い
```

です。

特に、

* 大規模既存業務システム
* 独自ルール
* 独自アーキテクチャ
* 独自運用
* 長期保守
* 調査中心

では、Skillの価値がかなり高いです。

---

# まず「Skillとは何か」

ここが重要です。

Prompt との違いが少し分かりづらいです。

---

# Prompt

```text id="d4os92"
「今回何をしてほしいか」
```

です。

例：

```text id="zjlwm1"
調査して
設計して
レビューして
```

---

# Skill

```text id="cjlwm2"
「この作業をするときの専門知識」
```

です。

つまり：

```text id="mjlwm3"
AIへ専門ノウハウを追加する
```

感じ。

---

# イメージ

例えばユーザーさん環境。

AIは普通：

```text id="ljlwm4"
一般的Web開発
```

寄りです。

でも実際は：

* WinForms
* .NET Framework 4.7
* SOAP
* レガシー
* Shift-JIS
* BOM
* AppLogger
* InstallShield
* Client/Server
* 既存互換性

など、

```text id="kjlwm5"
業務システム特有文化
```

があります。

これを教えるのが Skill。

---

# Promptとの違い

かなり重要。

---

# Prompt

```text id="jjlwm6"
一時的指示
```

---

# Skill

```text id="7jlwm7"
継続的専門知識
```

---

# 例えると

## Prompt

```text id="xjlwm8"
今回この調査して
```

---

## Skill

```text id="2jlwm9"
このシステムでは
Shift-JIS維持が重要
```

---

# ユーザーさん環境で特に有効なSkill

かなりあります。

---

# ① Legacy System Skill

超おすすめ。

```text id="jlwm10"
既存互換性優先
全面リファクタ禁止
既存コードスタイル優先
```

---

# ② WinForms Skill

かなり効きます。

```text id="jlwm11"
Designer.cs注意
Invoke注意
同期処理前提
```

---

# ③ DB Skill

かなり実務向き。

```text id="jlwm12"
StoredProcedure優先
Transaction維持
```

---

# ④ Logging Skill

ユーザーさんかなり合う。

```text id="jlwm13"
AppLogger利用
Catch時必ずログ
```

---

# ⑤ Feature Investigation Skill

これかなり強い。

```text id="jlwm14"
実装前に影響範囲調査
```

を習慣化できる。

---

# 実はSkillの本当の価値

超重要。

---

# 「思考パターン固定」

これです。

例えばAIは普通：

```text id="jlwm15"
まず実装
```

したがります。

でもSkillで：

```text id="jlwm16"
まず調査
影響範囲整理
```

を教え込める。

これはかなり大きい。

---

# Promptだけではダメ？

実は結構いけます。

ただ、

```text id="jlwm17"
毎回同じことを書く
```

問題があります。

例えば毎回：

```text id="jlwm18"
既存仕様維持
影響最小
Shift-JIS注意
```

を書く。

Skillにすると：

```text id="jlwm19"
常識化
```

できる。

---

# ただし注意

ここ重要。

---

# Skillは「重い」

Skillは：

```text id="jlwm20"
知識モジュール
```

なので、

大量に作ると：

* context圧迫
* relevance問題
* ノイズ化

しやすい。

---

# 実務おすすめ

かなり重要。

---

# Skillに向くもの

```text id="jlwm21"
長期的・普遍的ルール
```

---

# Promptに向くもの

```text id="jlwm22"
今回だけ
```

---

# つまり

## Skill

```text id="jlwm23"
この会社・このシステム文化
```

---

## Prompt

```text id="jlwm24"
今回何をやるか
```

---

# ユーザーさん環境ならおすすめSkill

かなり相性良い。

---

# ① legacy-system.skill.md

おすすめ。

内容：

```text id="jlwm25"
既存互換性
リファクタ禁止
小変更優先
```

---

# ② csharp-winforms.skill.md

おすすめ。

```text id="jlwm26"
WinForms文化
```

---

# ③ investigation.skill.md

かなりおすすめ。

```text id="jlwm27"
まず調査
影響範囲整理
```

---

# ④ installshield.skill.md

かなり強い。

ユーザーさん特有。

```text id="jlwm28"
ProductCode
UpgradeCode
共存インストール
```

---

# ⑤ client-server-compatibility.skill.md

かなり実務向き。

```text id="jlwm29"
後方互換性
Server変更最小
```

---

# かなり重要な結論

実務では：

---

# instructions

```text id="jlwm30"
軽いルール
```

---

# prompt

```text id="jlwm31"
今回の作業
```

---

# skill

```text id="jlwm32"
専門知識・思考様式
```

---

# agent

```text id="jlwm33"
人格・役割
```

この役割分担がかなり綺麗です。

---

# ユーザーさん環境でのおすすめ

かなりおすすめ。

最初は：

```text id="jlwm34"
Skillは2〜3個
```

くらいが良いです。

おすすめ：

```text id="jlwm35"
legacy-system.skill.md
investigation.skill.md
```

これだけでもかなり変わります。

