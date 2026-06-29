User:
では、複数の前提条件、例えば機能A、機能Bの追加があるとして、作業したい内容によってprompt.mdを用意して、必要に応じて　/prompt ~ というように質問時に読み込ませる、という感じの使い方になりますかね？


AI:
はい、まさにそのイメージです。
それがかなり実務向きの使い方です。

特にユーザーさんのような、

* 大規模既存システム
* 複数サブシステム
* 機能追加中心
* 調査 → 設計 → 実装 の流れがある

環境では非常に相性が良いです。

---

# イメージ

例えば：

```text id="h0o4c8"
機能A追加
機能B追加
```

があるとします。

---

# Promptを用途別に分ける

```text id="tqecdb"
.github/prompts/

  featureA-investigate.prompt.md
  featureA-design.prompt.md

  featureB-investigate.prompt.md
  featureB-design.prompt.md

  review-risk.prompt.md
```

みたいに作る。

---

# 作業時

例えば：

```text id="j1zqwx"
/prompt featureA-investigate
ログイン機能追加の影響範囲を調査してください
```

とか。

---

# 次

```text id="x88k78"
/prompt featureA-design
上記調査結果をもとに設計案を出してください
```

とか。

---

# さらに実装

```text id="mjlwmz"
/prompt implement-small-change
まずAuthServiceだけ修正してください
```

みたいに使う。

---

# 実は「機能別 + 作業別」がかなり強い

ここ重要です。

Promptには大きく2種類あります。

---

# ① 作業タイプPrompt

汎用。

例：

```text id="5n9tkx"
investigate.prompt.md
design.prompt.md
review.prompt.md
```

これは再利用性高い。

---

# ② 機能専用Prompt

今回だけ。

例：

```text id="0ezqg6"
featureA-login-extension.prompt.md
featureB-db-migration.prompt.md
```

---

# 実務おすすめは「混ぜる」

かなりおすすめ。

---

# ベースPrompt

汎用。

```text id="08g0q6"
調査だけ
設計だけ
実装だけ
```

など。

---

# 追加情報

今回仕様。

```text id="5pjlwm"
機能Aでは後方互換性重要
機能BではServer変更禁止
```

など。

---

# つまり

実務では：

```text id="yv06m3"
「Promptを組み合わせる」
```

感覚になります。

---

# かなり実務的な構成例

ユーザーさん環境ならかなりおすすめ。

```text id="r0wq6y"
.github/prompts/

  common/
    investigate.prompt.md
    design.prompt.md
    implement.prompt.md
    review.prompt.md

  features/
    feature-login-extension.prompt.md
    feature-visap-support.prompt.md
    feature-export-encryption.prompt.md
```

---

# 使い方イメージ

## 調査

```text id="mzmjlwm"
/prompt investigate
/prompt feature-visap-support

VisAP対応の影響範囲を調査してください
```

---

# 設計

```text id="6vz9qz"
/prompt design
/prompt feature-visap-support

影響を最小化する設計案を出してください
```

---

# 実装

```text id="0te2zk"
/prompt implement

Server側のみ最小変更で実装してください
```

---

# これの何が良いか

かなり重要。

---

# 「毎回長文書かなくてよい」

毎回：

```text id="o5ktg9"
既存仕様を維持して
影響範囲を調査して
後方互換性を保って
```

を書かなくてよくなる。

---

# 「AIの人格固定」

これかなり大きいです。

例えば：

## 調査Prompt

```text id="exn7yq"
コード変更禁止
```

を入れておく。

すると：

```text id="2mjlwm"
調査Agent化
```

しやすい。

---

# 「暴走防止」

AIは放置すると：

```text id="7rfwbv"
調査依頼
↓
勝手に実装
↓
勝手にリファクタ
```

しがち。

Prompt分離はかなり効きます。

---

# ただし注意点

ここ重要。

---

# Prompt増やしすぎ問題

実務で起きます。

```text id="j8zx6l"
featureA_v2_final_fix.prompt.md
```

みたいになり始める。

---

# おすすめ

Promptは：

```text id="zjlwmc"
作業単位
```

を中心にする。

---

# 良い

```text id="jlwm56"
investigate
design
implement
review
```

---

# 悪い

```text id="l88yl8"
login-button-left-side-change-v2.prompt.md
```

細かすぎる。

---

# 実務でかなりおすすめの考え方

重要。

---

# instructions

```text id="mnvsu7"
常時ルール
```

---

# prompt

```text id="jlwm98"
今回何をさせるか
```

---

# agent

```text id="2jlwm"
人格・専門性
```

---

# MCP

```text id="vzjlwm"
外部知識
```

これで役割分担するとかなり整理されます。
