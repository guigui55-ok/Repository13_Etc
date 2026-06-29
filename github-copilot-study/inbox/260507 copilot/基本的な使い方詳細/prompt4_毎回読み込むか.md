はい、ここはかなり重要な違いがあります。
結論から言うと：

```text id="g2f4rm"
prompt.md は、基本的に「自動では読み込まれない」
```

です。

つまり、

```text id="z7m1xw"
明示的に使う前提
```

と考えたほうがよいです。

---

# まず整理

## copilot-instructions.md

```text id="m8k3qv"
自動適用寄り
```

---

## *.instructions.md

```text id="v2n9py"
applyTo に応じて自動適用寄り
```

---

## skill.md

```text id="x6r1tb"
relevance に応じて半自動
```

---

## prompt.md

```text id="n5w7jc"
基本「手動」
```

これがかなり重要です。

---

# つまり

例えば：

```text id="d9p2fk"
.github/prompts/

  investigate.prompt.md
  design.prompt.md
  implement.prompt.md
```

があっても、

普通に：

```text id="m4t8zy"
この機能を追加してください
```

と聞いただけでは、

```text id="r7k1vs"
自動で investigate.prompt.md が適用
```

されたりは基本しません。

---

# Prompt File の本質

Prompt File は：

```text id="b3q9xf"
「保存済み定型プロンプト」
```

です。

つまり：

```text id="v8n4jk"
毎回コピペしなくて済む
```

ためのもの。

---

# なので

例えば：

```text id="f6m2wp"
/prompt investigate
```

を明示的に使う。

---

# これがかなり重要

つまり Prompt は：

```text id="h1r5zk"
AIの自動動作を変えるもの
```

ではなく、

```text id="m7q4vx"
ユーザーが呼び出す作業テンプレート
```

です。

---

# なので安心

これは逆にメリットです。

もし自動だったら：

```text id="x4k8mj"
関係ないprompt全部混ざる
```

危険があります。

---

# 例えば

```text id="p2m7qw"
featureA-design.prompt.md
featureB-design.prompt.md
```

がある。

---

# 自動適用だったら

AIが：

```text id="k5t1vx"
両方読んで混乱
```

する可能性がある。

---

# Promptが手動なのは合理的

つまり：

```text id="u7r3mb"
今回必要なものだけ使う
```

設計です。

---

# 実務でかなり重要

おすすめ運用。

---

# 普段

```text id="d4m9zp"
instructions
skill
```

で、

基本文化・知識を固定。

---

# 作業時だけ

```text id="w1k7qy"
/prompt investigate
```

```text id="f9r2mx"
/prompt design
```

を使う。

---

# かなり良い理解

つまり：

---

# instructions

```text id="v6q1tw"
常時ON
```

---

# skill

```text id="z8m4pk"
半自動ON
```

---

# prompt

```text id="x3r7vy"
必要時だけON
```

です。

---

# ユーザーさん環境だとかなり重要

特に大規模Workspaceでは：

```text id="q7m2fk"
不要prompt自動混入
```

は危険です。

なので：

```text id="h5n8qv"
Promptは手動
```

のほうが実務的に安全です。

---

# 実務でおすすめ

かなりおすすめ。

---

# instructions

少数・軽量。

---

# skill

少数・専門特化。

---

# prompt

大量OK。

必要時だけ呼ぶ。

---

# これが一番スケールする

例えば：

```text id="p4r9mk"
prompts/
```

に、

50個

あっても問題ない。

理由：

```text id="t8v2qx"
毎回全部読まない
```

からです。

---

# 最後にかなり重要

実は Prompt File は、

```text id="n6m3zk"
「作業マクロ」
```

みたいな感覚です。

つまり：

```text id="g1q5vr"
/prompt investigate
```

すると、

毎回同じ調査方針を
瞬時に呼べる。

これはかなり強いです。
