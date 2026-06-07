User:
skill.mdに apply to のようなもので指定可能ですか？

AI:
現状の考え方としては、

```text id="cq3y8w"
.instructions.md の applyTo のような
「公式なファイルパターン適用」
```

は、Skillには基本ありません。

ここかなり重要です。

---

# つまり

## instructions

```yaml id="8m0f7p"
applyTo: "**/*.cs"
```

のように：

```text id="f4y1kq"
ファイルパターン適用
```

できます。

---

# でも Skill

Skill は：

```text id="g7u2vx"
「専門知識モジュール」
```

なので、

```text id="u9w3zs"
applyTo 的な公式仕組みは薄い
```

です。

---

# なぜか

役割が違うためです。

---

# instructions

```text id="7k8m2r"
ファイルベース
```

---

# skill

```text id="9x5p1n"
タスク・知識ベース
```

だから。

---

# Skill の relevance は何で決まる？

ざっくり：

* Skill名
* Skill説明
* Prompt内容
* 現在作業
* 開いているコード
* Agent
* MCP情報

などです。

つまり：

```text id="5m9q7x"
意味的 relevance
```

で選ばれる。

---

# 例えば

## winforms.skill.md

内容：

```text id="m3t9vq"
WinForms
Invoke
Designer.cs
```

---

# 今回

```text id="n7k2fp"
Form追加したい
```

なら：

```text id="2r6v1z"
WinForms Skill relevance 高
```

になります。

---

# installshield.skill.md

内容：

```text id="x8w4bc"
UpgradeCode
ProductCode
```

なら：

```text id="w5m2jh"
Installer作業時だけ relevance 高
```

になる。

---

# 実務でかなり重要

つまり Skill は：

```text id="m8x4vf"
「名前と内容で relevance を誘導」
```

します。

---

# なのでおすすめ

かなり重要。

---

# Skill名を専門特化する

悪い：

```text id="v2y9kc"
common.skill.md
```

---

良い：

```text id="p4r7zw"
winforms-ui.skill.md
```

```text id="s9n1xm"
oracle-transaction.skill.md
```

```text id="d6k5qh"
installshield-upgrade.skill.md
```

---

# 内容も専門化

悪い：

```text id="u3m8pv"
会社ルール全部
```

---

良い：

```text id="q5z1bn"
WinForms限定知識
```

---

# ただし

# 擬似 applyTo はできる

かなり実務的なテクです。

---

# Skill冒頭に対象を書く

例えば：

```md id="m7c2qt"
# 対象

- ServerApp/**/*.cs
- SOAP通信
- Client/Server互換性
```

---

# または

```md id="v8p1xr"
# このSkillを使う場面

- ServerApp配下の作業
- 通信IF変更
- 後方互換性維持
```

---

# こうすると

AI側が：

```text id="t4n6my"
「あ、このSkill関係ありそう」
```

と判断しやすくなる。

---

# 実はAgentと組み合わせると強い

かなりおすすめ。

---

# 例

## server.agent.md

```text id="d7m9pk"
server関連作業時は
server-compatibility.skill.md
を優先利用
```

---

これで：

```text id="h1w5qr"
擬似 applyTo
```

っぽくできます。

---

# 実務でおすすめ構成

ユーザーさん環境ならかなり良い。

---

# instructions

```text id="f6n2zt"
ファイルベース制御
```

例：

* *.cs
* DB/**/*.sql
* *Form*.cs

---

# skills

```text id="r3m8xq"
専門知識
```

例：

* legacy
* installshield
* compatibility
* winforms-threading

---

# prompts

```text id="b5k7nv"
今回作業
```

---

# agents

```text id="g2q9xf"
作業人格
```

---

# 一番おすすめ

かなり重要。

Skillに：

```text id="k9w4zc"
対象システム
対象ディレクトリ
使う場面
```

を書く。

これが実質：

```text id="x1r6mt"
擬似 applyTo
```

としてかなり効きます。

---

# ユーザーさん環境でかなり有効そうなSkill例

例えば：

```text id="t7m3pq"
skills/

  legacy-system.skill.md
  winforms-threading.skill.md
  installshield-upgrade.skill.md
  oracle-transaction.skill.md
  client-server-compatibility.skill.md
```

これはかなり相性良いです。
