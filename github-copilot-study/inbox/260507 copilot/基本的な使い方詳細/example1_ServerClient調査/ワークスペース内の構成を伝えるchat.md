User:
github Copilot の instructions や カスタムagent について、 ワークスペース内に複数のプログラム（インストーラー、サーバー側、メインエディションクライアント、別エディションクライアント、ツールなど）が混在する場合、この情報も記載しておいた方がよいですか？ 例） ●●フォルダ：サーバー △△フォルダ：インストーラー関連

Chat:
はい、そのケースは **instructions / カスタムAgent に記載する価値がかなり高い** です。
特に、**ワークスペース内に複数プロダクト・複数責務が混在している場合は、フォルダ構成と責務の説明を書いておくと誤編集・誤参照をかなり減らせます。**

ただし、全部を詳細に書くのではなく、**「AIが判断に必要なレベル」だけ書く**のがおすすめです。

おすすめ順としては、

1. **フォルダ → 役割（必須）**
2. **依存関係（推奨）**
3. **変更禁止ルール（必要なら）**
4. **ビルド・実行方法（必要なら）**

です。

例えばこういう感じです。

```md
# Workspace Overview

このワークスペースには複数アプリが含まれる。

## Components

- Server/
  - サーバーアプリ
  - API提供
  - DBアクセスあり

- Installer/
  - インストーラー生成
  - 配布用パッケージ管理
  - 通常の機能修正では編集しない

- ClientMain/
  - メインエディション（フル機能）

- ClientLite/
  - 別エディション（制限版）
  - ClientMain と共通ロジックあり

- Tools/
  - 開発支援ツール
  - 本番コードではない

## Dependency

ClientMain
  → Server API

ClientLite
  → Server API

Installer
  → ClientMain
  → ClientLite

## Rules

- 修正対象が指定されていない場合、編集前に対象コンポーネントを確認する
- Installer は必要時のみ変更する
- 共通処理変更時は ClientMain / ClientLite 両方への影響確認を行う
```

これを書くと、例えば

「ログ出力追加して」

だけ依頼したときでも、

❌ Installer側まで変更
❌ Lite版にも勝手に反映
❌ ServerとClient両方変更

みたいな事故が減ります。

さらに案件規模が大きいなら、以前話していた **Agent + Skill 分離** と相性が良いです。

例：

```text
AGENT.md
├─ workspace構成
├─ 用語集
├─ 修正判断ルール

skills/
├─ modify-server/
├─ modify-installer/
├─ release-check/
├─ investigate-api/
```

すると、

* Agent → 「どこを見るか」
* Skill → 「どう作業するか」

に責務分離できます。

一点だけ注意すると、**フォルダ一覧を生で全部書くのは逆効果**です。

避けたい例：

```md
src/
src2/
src_old/
src_backup/
tmp/
tmp2/
tool_old/
tool_new/
...
```

こうするとコンテキストだけ増えて判断精度が落ちます。

代わりに、

```md
MainProduct/*
Installer/*
Tools/*
Legacy/*
```

のように**論理単位（役割単位）でまとめる**方が長期運用しやすいです。

今回の「サーバー・複数クライアント・インストーラー混在」は、かなり典型的に instructions に向いているケースです。


# -------------
User:
instructionsに書いて毎回読み込まれると、コンテキスト増加を懸念していますが、いかがでしょうか？

Chat:
その懸念はかなり正しいです。
実際、**instructions（copilot-instructions.md / AGENT.md）に何でも書くと、精度向上より先にコンテキスト汚染・誤誘導が起きる**ことがあります。

結論から言うと、

* **毎回必要な情報 → instructions**
* **必要時だけ使う情報 → Skill**
* **作業ごとの条件 → Prompt**
* **詳細資料 → 別ファイル参照**

に分離するのが基本です。

特に、以前話していたような「複数案件・複数コンポーネント構成」の場合、instructionsはかなり薄くした方が安定します。

おすすめはこのくらいです。

### 良い（薄い instructions）

```md
# Workspace Overview

複数コンポーネントを含む。

- Server/ : API・DB
- Installer/ : 配布生成
- ClientMain/ : 標準版
- ClientLite/ : 制限版
- Tools/ : 開発支援

ルール:
- 対象が不明なら確認
- 複数コンポーネントへ変更しない
- 共通変更は影響範囲確認
```

→ 200～500文字程度

これは毎回読まれても許容範囲です。

---

### 悪い（厚い instructions）

```md
Installer:
・setup.exe生成
・WiX
・署名
・配布手順
・例外処理

Server:
・API一覧
・URL
・認証
・DB設計
…
```

数千文字。

→ 毎回不要な知識が入る。

---

むしろ、詳細は別ファイル化します。

```text
.github/
├─ instructions.md ← 最小
├─ architecture/
│   ├─ system.md
│   ├─ installer.md
│   ├─ server.md
│   └─ client.md
└─ skills/
    ├─ modify-server/
    └─ release-build/
```

instructions：

```md
詳細は参照：
- architecture/server.md
- architecture/installer.md
必要時のみ読む
```

こうしておく。

---

もう一段踏み込むと、**「毎回読む情報」と「呼ばれた時だけ読む情報」を意識して階層化**するとよいです。

目安：

| 種類           |  毎回読む |      推奨サイズ |
| ------------ | ----: | ---------: |
| instructions |     ○ | 300～1000文字 |
| AGENT.md     |     △ | 500～2000文字 |
| SKILL.md     | 必要時のみ |      1～5KB |
| 補助md         | 必要時のみ |       制限緩め |
| prompt.md    | 明示指定時 |       制限緩め |

あなたが前に懸念していた「案件増加でSkillが増えて誤使用しそう」という話ともつながりますが、**instructionsに案件一覧を書くより、案件→Skillへ逃がす方がスケールしやすい**です。

なので今回の例なら、

❌ instructions に
「Serverの仕様」「Installerの手順」「Client差分」を全部書く

ではなく、

⭕ instructions に
「どのフォルダが何か」「変更時の原則」だけ書く

くらいがバランス良いです。

