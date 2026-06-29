User:
作業単位にすると、promptはかなり多くなりそうなイメージです。
フォルダ分けできたりしますかね？ /prompt/featureA,  /prompt/feartuerB, ... など

AI:
はい、できます。
そして実務では、

```text id="g6q2rm"
「フォルダ分け前提」
```

くらいで考えたほうがよいです。

特にユーザーさんのような、

* 大規模Workspace
* 複数システム
* 長期保守
* 機能追加多数

では、Promptは普通に増えます。

なので、

```text id="k3i9g8"
promptをディレクトリ管理する
```

のはかなりおすすめです。

---

# まず結論

例えば：

```text id="hbl89g"
.github/prompts/

  common/
    investigate.prompt.md
    design.prompt.md
    implement.prompt.md
    review.prompt.md

  featureA/
    investigate-impact.prompt.md
    design-extension.prompt.md
    review-risk.prompt.md

  featureB/
    investigate-db.prompt.md
    design-server.prompt.md

  visap/
    compatibility-check.prompt.md
```

みたいな構成でOKです。

---

# 呼び出しイメージ

実際のUI/CLIは時期で多少変わりますが、
概念としては：

```text id="trm2nm"
/prompt featureA/investigate-impact
```

みたいな感覚になります。

つまり、

```text id="o4mxna"
パス指定っぽく扱える
```

イメージです。

---

# 実務でかなり重要

ここ重要です。

---

# Promptは「資産化」する

最初は：

```text id="jlwmq1"
数個
```

ですが、

実務ではすぐ：

```text id="jlwmq2"
20〜50
```

くらいになります。

---

# だから分類が必要

特に：

* 調査
* 設計
* レビュー
* DB
* Server
* WinForms
* Test
* Security

などが混ざる。

---

# おすすめ構成

ユーザーさん環境ならかなりおすすめ。

---

# ① common

全案件共通。

```text id="jlwmq3"
common/
```

例：

```text id="jlwmq4"
investigate.prompt.md
design.prompt.md
implement.prompt.md
review.prompt.md
```

---

# ② feature

機能別。

```text id="jlwmq5"
feature-export/
feature-login/
feature-visap/
```

---

# ③ domain

技術領域。

```text id="jlwmq6"
database/
server/
winforms/
```

---

# ④ temporary

一時作業。

```text id="jlwmq7"
tmp/
```

ここかなり重要。

調査終わったら削除。

---

# 実は feature単位だけだと辛い

かなり実務的な話です。

例えば：

```text id="jlwmq8"
featureA/
  investigate.prompt.md
  design.prompt.md
```

だけにすると、

似たPromptが大量増殖します。

---

# おすすめは「共通 + 差分」

かなりおすすめ。

---

# 共通Prompt

```text id="jlwmq9"
common/investigate.prompt.md
```

---

# feature側

```text id="jlwmqa"
featureA/context.prompt.md
```

中身：

```text id="jlwmqb"
- Login系機能
- Server変更禁止
- DB変更禁止
- 後方互換性重要
```

---

# 呼び出し

```text id="jlwmqc"
/prompt common/investigate
/prompt featureA/context
```

こういう構成がかなり強いです。

---

# これの何が良いか

超重要。

---

# Prompt重複が減る

悪い例：

```text id="jlwmqd"
featureA-investigate
featureB-investigate
featureC-investigate
```

全部似る。

---

# 共通化すると

```text id="jlwmqe"
investigate
```

は1個だけ。

---

# featureは差分だけ

これがスケールします。

---

# 大規模開発でおすすめ構成

かなりおすすめ。

```text id="jlwmqf"
.github/prompts/

  common/
    investigate.prompt.md
    design.prompt.md
    implement.prompt.md
    review.prompt.md

  domains/
    database.prompt.md
    server.prompt.md
    winforms.prompt.md
    legacy.prompt.md

  features/
    feature-login/
      context.prompt.md

    feature-visap/
      context.prompt.md

    feature-export-encryption/
      context.prompt.md
```

---

# 実際の運用イメージ

例えば：

## 調査

```text id="jlwmqg"
/prompt common/investigate
/prompt domains/server
/prompt features/feature-visap/context
```

---

# 設計

```text id="jlwmqh"
/prompt common/design
/prompt domains/legacy
/prompt features/feature-visap/context
```

---

# かなり重要

実はこの構成、

```text id="jlwmqi"
AIに「必要最小限だけ読ませる」
```

という意味でも良いです。

つまり：

```text id="jlwmqj"
relevance最適化
```

になっています。

---

# 最後にかなり重要な話

Promptを整理し始めると、

実質：

```text id="jlwmqk"
「AI向け設計資産」
```

になります。

つまり、

* 調査観点
* 設計観点
* レビュー観点
* 保守方針

が蓄積される。

これはかなり価値あります。
