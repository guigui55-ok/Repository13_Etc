User:
csharp.instructions.md について、冒頭に applyTo: "**/*.cs" がありますが、これは必須ですか？また、csファイルを扱うとき必ず読み込まれるということですか？

AI:
はい、その `applyTo` はかなり重要です。

まず結論から言うと：

```text id="bqv0i7"
.instructions.md では、applyTo は実質ほぼ必須
```

と思ってよいです。

そして、

```text id="ztj3wu"
applyTo に一致したファイルをCopilotが扱う時、
その instructions が「候補として適用される」
```

というイメージです。

ただし、

```text id="c7gm4s"
「絶対100%毎回読み込まれる」
```

保証ではありません。

（内部的には relevance / context window / task relevance で選別される）

---

# まず `.instructions.md` の役割

これは、

```text id="3zps59"
「特定ファイル群向けの追加ルール」
```

です。

なので：

```md id="7q0l1n"
---
applyTo: "**/*.cs"
---
```

を書くことで、

```text id="fxgpxm"
「C#ファイル扱う時にこの指示を使ってね」
```

になります。

---

# applyTo が無いと？

挙動が曖昧になります。

Copilot側が：

```text id="74uv0k"
「これは何に適用するinstructions？」
```

を判断しづらくなる。

結果：

* 効かなかったり
* relevance低下したり
* 適用優先度が落ちたり

しやすいです。

なので基本は書いたほうがよいです。

---

# applyTo は glob

これは glob pattern です。

例えば：

## C#

```yaml id="x4m5cr"
applyTo: "**/*.cs"
```

---

## WinFormsだけ

```yaml id="qucn3d"
applyTo: "**/*Form*.cs"
```

---

## DB関連だけ

```yaml id="7gcr2r"
applyTo: "DB/**/*.sql"
```

---

## Serverだけ

```yaml id="s3bgfr"
applyTo: "ServerApp/**/*.cs"
```

---

## 複数

```yaml id="wyvwnb"
applyTo:
  - "ClientApp/**/*.cs"
  - "SharedLib/**/*.cs"
```

---

# 「必ず読み込まれるのか？」

ここが重要です。

実際は：

```text id="r52tdq"
完全強制ではない
```

です。

Copilot内部では、

* relevance
* 現在開いているファイル
* 編集対象
* プロンプト内容
* context長

などで選別されます。

つまり：

```text id="h4lz1g"
「対象なら優先的に参照候補になる」
```

くらいの理解が近いです。

---

# 実務感覚だとかなり効く

例えば：

```md id="wn71lz"
applyTo: "**/*.cs"
```

で、

```md id="q0nl0t"
async禁止
```

を書いておくと、

かなり async 提案が減ります。

つまり効いています。

---

# ただし「強すぎる applyTo」は危険

例えば：

```yaml id="cw4lsw"
applyTo: "**/*"
```

に大量instructionsを書くと、

全部の作業で毎回参照候補になり、

* コンテキスト圧迫
* relevance低下
* ノイズ化

しやすいです。

---

# 実務おすすめ

かなり重要です。

## ① 共通は最小

```text id="c1syf8"
copilot-instructions.md
```

は短く。

---

## ② instructions は狭く

これ重要。

悪い例：

```yaml id="j9s71e"
applyTo: "**/*.cs"
```

だけで巨大instructions。

---

良い例：

```yaml id="l7u1x4"
applyTo: "ServerApp/**/*.cs"
```

```yaml id="jl7b2i"
applyTo: "**/*Form*.cs"
```

```yaml id="aq6x4k"
applyTo: "DB/**/*.sql"
```

みたいに分ける。

---

# ユーザーさん環境ならおすすめ構成

かなり相性良いです。

```text id="g8y0vh"
.github/instructions/

  legacy.instructions.md
    → 全C#

  winforms.instructions.md
    → *Form*.cs

  database.instructions.md
    → DB/**/*.sql

  server.instructions.md
    → Server/**/*.cs

  test.instructions.md
    → Tests/**/*.cs
```

---

# 特に重要

実は：

```text id="f8qlb0"
instructionsを細かく分ける
```

ほうが、

```text id="8x8odk"
長大1ファイル
```

より効きやすいです。

理由：

Copilot側が relevance 判定しやすいため。

---

# 最後に重要な点

`.instructions.md` は、

```text id="m80qba"
「絶対命令」
```

ではなく、

```text id="qh3esq"
「強いヒント」
```

です。

なので本当に重要なことは、

毎回Prompt側にも書くのが安全です。

例えば：

```text id="y28bif"
コード変更せず調査だけしてください
```

これは毎回Promptに書くべきです。
