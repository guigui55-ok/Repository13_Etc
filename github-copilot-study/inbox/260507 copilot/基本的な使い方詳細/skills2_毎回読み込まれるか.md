User:
複数のskill.mdがあるとして、毎回すべてのファイルが読み込まれますか？


AI:
これ、かなり重要なポイントです。
結論から言うと：

```text id="1q4d5u"
「毎回すべて強制的に全文読み込み」
ではありません。
```

です。

実際は、

```text id="5t4j9g"
「今回の作業に関係ありそうな Skill が選ばれる」
```

イメージです。

ただし、

```text id="z6s0z2"
Skillの作り方が悪いと、
全部 relevance 高扱いになってしまう
```

ことがあります。

---

# まず大前提

AIには：

```text id="x2j1fw"
context window（読める量）
```

制限があります。

なので内部では：

```text id="t8sl1v"
「必要そうな情報だけ選ぶ」
```

処理をしています。

Skillもその対象です。

---

# イメージ

例えば：

```text id="q9s8me"
skills/

  winforms.skill.md
  database.skill.md
  installshield.skill.md
  legacy.skill.md
```

があるとします。

---

# 今回の作業

```text id="dxjv8h"
WinForms画面追加
```

なら：

---

# relevance 高

```text id="yy9cxk"
winforms.skill.md
legacy.skill.md
```

---

# relevance 低

```text id="4r8c5w"
installshield.skill.md
database.skill.md
```

---

# つまり

内部的には：

```text id="h7zj9w"
「必要そうなSkillだけ優先投入」
```

されます。

---

# ただし重要

# 「全部読まれる状態」は普通に起きる

これはあります。

---

# 悪い例

例えば：

## legacy.skill.md

```text id="j2f7mu"
全システム共通ルール
```

---

## database.skill.md

```text id="gm8v2n"
全システム共通ルール
```

---

## winforms.skill.md

```text id="w6s9pe"
全システム共通ルール
```

全部似ている。

すると：

```text id="4m7n0s"
全部 relevance 高
```

になりやすい。

---

# だから重要なのは

```text id="l5q1hz"
Skillの責務分離
```

です。

---

# 実務おすすめ

かなり重要。

---

# 良いSkill

```text id="u9w4cx"
狭く・専門的
```

---

# 悪いSkill

```text id="0k8m1a"
巨大・汎用・何でも入り
```

---

# 良い例

## winforms.skill.md

```text id="5y2d4z"
Invoke
Designer.cs
UIスレッド
```

専用。

---

## installshield.skill.md

```text id="n4c9xe"
ProductCode
UpgradeCode
共存インストール
```

専用。

---

# 悪い例

```text id="x3j9ks"
会社ルール全部入り
```

これやると：

```text id="q0w2zr"
毎回 relevance 高
```

になりやすい。

---

# Promptとの大きな違い

重要。

---

# Prompt

```text id="f5n7vb"
手動選択
```

---

# Skill

```text id="k8x1mu"
自動選択寄り
```

です。

---

# だからSkillは慎重に作る

かなり重要。

---

# 実務でおすすめ

ユーザーさん環境なら：

---

# Skillは少数精鋭

おすすめ：

```text id="z7p1qj"
3〜6個程度
```

---

# そして専門特化

例：

```text id="w2m5nc"
legacy-system.skill.md
```

↓

```text id="m8r4tx"
既存互換性
```

だけ。

---

# installshield.skill.md

↓

```text id="z3j9uy"
Installer特殊知識
```

だけ。

---

# Prompt側へ逃がす

重要。

例えば：

```text id="y1s6ko"
今回のFeature条件
```

はSkillに入れない。

Promptへ。

---

# ユーザーさん環境でおすすめ構成

かなり現実的。

---

# Skill

```text id="m7q9pf"
長期知識
```

例：

* legacy
* winforms
* installshield
* client-server-compatibility

---

# Prompt

```text id="r4x2dw"
今回作業
```

例：

* 調査
* 設計
* 機能A
* 機能B

---

# instructions

```text id="h9v1ek"
軽量ルール
```

---

# 一番重要な理解

Skillは：

```text id="m2t8qy"
「専門知識キャッシュ」
```

みたいなものです。

なので：

```text id="b1y7cf"
大量に作る
↓
全部汎用
↓
全部 relevance 高
↓
context圧迫
```

が一番危険です。

---

# 実務感覚だと

かなりおすすめなのは：

```text id="s4k9mw"
Skillは少なく
Promptは多く
```

です。

これがスケールしやすいです。

