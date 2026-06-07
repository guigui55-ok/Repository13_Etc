User:
prompt.mdの使い方について investigate-feature.prompt.md というのは固定名ですか？（任意に名前を付けられますか？） また、複数のprompr.mdを設置できますか？（できるなら、切り替え方法は？）

AI:
はい、かなり自由です。
そして、ここは重要なのですが、

```text id="5n1j1n"
Prompt Files は「固定名」ではありません。
```

完全に任意名です。

---

# まず結論

## investigate-feature.prompt.md は固定名？

```text id="2ocw0c"
固定名ではない
```

です。

例えば：

```text id="2r4n5w"
investigate-feature.prompt.md
```

は単なるサンプル名です。

---

# 任意名OK

例えば：

```text id="zyh1fd"
.github/prompts/

  調査用.prompt.md
  設計検討.prompt.md
  DB影響調査.prompt.md
  WinForms追加.prompt.md
  SQLレビュー.prompt.md
```

でもOKです。

英語でなくても大丈夫です。

---

# 実務では名前超重要

かなり重要です。

おすすめ：

```text id="x64a8w"
何をするPromptか
```

をファイル名で明確化。

---

# 良い例

```text id="u3px7t"
investigate-feature.prompt.md
```

```text id="glv9e9"
review-db-impact.prompt.md
```

```text id="j1xgs4"
design-server-extension.prompt.md
```

```text id="j3jzjg"
winforms-ui-addition.prompt.md
```

---

# 悪い例

```text id="aqm89k"
prompt1.md
```

```text id="y0cn0h"
test.md
```

あとで分からなくなる。

---

# 複数設置できますか？

```text id="lzmd7u"
はい、できます。
```

むしろ、

```text id="oc0m2d"
複数作る前提
```

です。

---

# 実務では大量に増える

例えば：

```text id="pn1q6o"
.github/prompts/

  investigate-feature.prompt.md
  investigate-db-impact.prompt.md
  design-feature.prompt.md
  implement-small-change.prompt.md
  review-security.prompt.md
  generate-testcases.prompt.md
  summarize-legacy-flow.prompt.md
  migrate-config.prompt.md
```

みたいになります。

---

# 一番重要

# 「切り替え方法」

ここ重要です。

現在のGitHub Copilot系では、

```text id="3d4n8l"
Prompt File を「明示的に選んで使う」
```

形です。

つまり：

```text id="m34z4q"
自動切替ではない
```

です。

---

# イメージ

VSCodeやCopilot Chatで：

```text id="8rm1jz"
/prompts: investigate-feature
```

みたいに選ぶ。

またはGUIから選択。

（UIは時期で結構変わる）

---

# つまり

Prompt Files は：

```text id="3uhr5u"
「再利用可能な定型プロンプト」
```

です。

---

# instructionsとの違い

ここ超重要です。

## instructions

```text id="kv3f5k"
自動適用寄り
```

---

## prompt.md

```text id="uvqeb8"
手動呼び出し
```

です。

---

# 実務での使い分け

かなりおすすめ。

---

# instructions

「常時ルール」

例：

```text id="2t2y98"
既存仕様維持
async禁止
AppLogger使用
```

---

# prompt.md

「今回何をさせるか」

例：

```text id="d5x7zv"
調査だけ
```

```text id="hccv5f"
設計だけ
```

```text id="75p7c5"
テスト観点出力
```

---

# ユーザーさん環境でおすすめ

かなり合います。

---

# ① 調査系

```text id="jlwmrf"
investigate-feature.prompt.md
investigate-db-impact.prompt.md
trace-processing-flow.prompt.md
```

---

# ② 設計系

```text id="xsvmz4"
design-feature.prompt.md
design-minimal-impact.prompt.md
```

---

# ③ 実装系

```text id="lkjlwm"
implement-small-change.prompt.md
implement-safe-refactor.prompt.md
```

---

# ④ レビュー系

```text id="d8a75m"
review-risk.prompt.md
review-legacy-impact.prompt.md
```

---

# 実はかなり便利な運用

かなりおすすめ。

例えば：

## Prompt

```text id="61l3c0"
investigate-feature.prompt.md
```

で調査。

↓

## 次

```text id="n87tkj"
design-feature.prompt.md
```

で設計。

↓

## 次

```text id="xlyb5y"
implement-small-change.prompt.md
```

で実装。

みたいに、

```text id="ayb0vb"
工程ごとにAI人格を切り替える
```

感覚で使えます。

---

# 実務ではこれかなり効く

AI暴走を減らせます。

特に：

```text id="7bry2u"
「調査だけして」
```

をPrompt化しておくのは非常に有効です。

これをやらないと、
かなり高確率で実装を始めます。
