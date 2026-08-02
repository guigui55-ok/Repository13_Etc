User:
ローカルではVsCodeで作業しますが、どのようなフォルダ構成にしてたら良いか教えてください。 あと、上記（ChatGPTのPrjectでの作業）の具体的な手順を教えてください


ChatGPT:
はい。
Toshiさんの場合は、**ローカルのVS Code側を正本（Source of Truth）にして、ChatGPT Projectは「作業・レビュー・検討の場」として使う**構成をおすすめします。

ChatGPT Projectは、チャット・ファイル・プロジェクト指示を1か所にまとめ、継続的な作業に利用できます。プロジェクト内にはファイルを追加でき、プロジェクト専用指示も設定できます。([OpenAI Help Center][1])

私なら、次のように始めます。

---

# 1. おすすめのローカルフォルダ構成

今回の用途なら、まずは以下くらいがよいと思います。

```text
software-project/
│
├── .github/
│   ├── agents/
│   │   ├── doc-sync.agent.md
│   │   └── copilot-config-reviewer.agent.md
│   │
│   └── copilot-instructions.md
│
├── docs/
│   ├── 00_change_requests/
│   │   ├── CR-0001_json-to-xml.md
│   │   └── ...
│   │
│   ├── 01_research/
│   │   └── research_report.md
│   │
│   ├── 02_estimate/
│   │   └── estimate.md
│   │
│   ├── 03_requirements/
│   │   └── requirements.md
│   │
│   ├── 04_basic_design/
│   │   └── basic_design.md
│   │
│   ├── 05_function_spec/
│   │   └── function_spec.md
│   │
│   ├── 06_detail_design/
│   │   └── detail_design.md
│   │
│   ├── 07_test_design/
│   │   └── test_design.md
│   │
│   ├── 08_test_spec/
│   │   └── test_spec.md
│   │
│   ├── 09_operation_design/
│   │   └── operation_design.md
│   │
│   └── 10_operation_manual/
│       └── operation_manual.md
│
├── chatgpt/
│   ├── project-instructions.md
│   │
│   ├── workflows/
│   │   ├── impact-analysis.md
│   │   ├── doc-sync.md
│   │   └── config-review.md
│   │
│   └── logs/
│       └── decisions.md
│
├── src/
│   └── ...
│
├── tests/
│   └── ...
│
└── README.md
```

## ポイントは `chatgpt/` を設けることです

私は今回、これをおすすめします。

```text
chatgpt/
├── project-instructions.md
├── workflows/
└── logs/
```

理由は、ChatGPTのWeb画面だけに設定内容を置くと、

* 何を設定したかGitで履歴管理しにくい
* 修正前後を比較しにくい
* 別Projectへ転用しにくい
* VS Code側から内容をレビューできない

ためです。

つまり、

```text
ローカル
chatgpt/project-instructions.md
        ↓
内容をChatGPT Project instructionsへ設定
```

という運用です。

**ChatGPT画面上の設定を正本にせず、ローカルのMarkdownを正本にする**方がToshiさんの現在のGit・VS Code中心の運用には合っていると思います。

---

# 2. GitHub Copilot用とChatGPT用は分けるべきか

私は、**ファイル自体は分けるが、考え方は共有する**のがよいと思います。

例えば、

```text
.github/agents/doc-sync.agent.md
```

と

```text
chatgpt/workflows/doc-sync.md
```

です。

完全に同一内容をコピーするのではなく、

```text
doc-sync.agent.md
    ↓
GitHub Copilotが
リポジトリを直接検索・編集するための指示


chatgpt/workflows/doc-sync.md
    ↓
ChatGPTで
調査・レビュー・修正案作成を行うための指示
```

と少し役割を変えます。

ChatGPTとVS Code Copilotでは、操作できる環境やワークフローが違うためです。

---

# 3. ローカルとChatGPT Projectの役割分担

私は次のように分けます。

## VS Code + GitHub Copilot

主に、

```text
実ファイルの検索
↓
影響調査
↓
複数ファイルの直接修正
↓
Git diff確認
↓
commit
```

を担当します。

## ChatGPT Project

主に、

```text
変更方針の検討
↓
広い視点での影響調査
↓
ドキュメント間の整合性レビュー
↓
修正方針の検討
↓
Agent構成のレビュー
```

を担当させます。

私としては、両方に完全に同じことをさせるより、

> **Copilot：実作業に強い**
>
> **ChatGPT Project：検討・レビュー・継続的な文脈管理に使う**

という使い分けから始めるのがよいと思います。

---

# 4. ChatGPT Projectの具体的な作成手順

ここから実際の作業手順です。

## STEP 1：新しいProjectを作る

ChatGPTブラウザ版のサイドバーから、

```text
新規プロジェクト
```

を選択します。

その後、

```text
名前:
Software Development Documents
```

または、

```text
ソフトウェア開発ドキュメント
```

などを設定します。

公式の手順でも、サイドバーの「新規プロジェクト」から作成し、名前・アイコン・色を設定する流れになっています。([OpenAI Help Center][1])

私は日本語で、

```text
開発ドキュメント管理
```

くらいでもよいと思います。

---

## STEP 2：メモリ設定を選ぶ

新規Project作成時に選択できるなら、今回の用途では私は、

```text
Project-only memory
```

をおすすめします。

Project-only memoryでは、そのProject内に文脈を閉じた運用ができます。公式ヘルプによると、Project作成時にDefault memoryまたはProject-only memoryを選択でき、既存Projectを後からProject-onlyへ切り替えることはできません。([OpenAI Help Center][2])

ソフトウェア案件単位で使う場合、

```text
案件A
案件B
個人学習
```

の情報を混ぜない方が自然だからです。

---

# 5. Project instructionsを設定する

Project右上の3点メニューから、

```text
Project settings
```

を開き、Project instructionsを設定します。

この指示はそのProject内だけに適用され、グローバルなCustom Instructionsより優先されます。([OpenAI Help Center][1])

ローカルには、

```text
chatgpt/project-instructions.md
```

として同じ内容を保存します。

最初は次くらいでよいと思います。

# プロジェクトの目的

本プロジェクトでは、ソフトウェア開発に関する調査、要件、設計、実装、テスト、運用の検討と、関連ドキュメントの作成・レビュー・変更反映を行う。

# 基本方針

* 要件、設計、仕様、実装、テスト、運用の整合性を重視する。
* 変更内容を単純な文字列置換として扱わず、各記述の文脈とドキュメントの目的を確認する。
* 変更履歴、調査経緯、不採用理由など、残すべき情報を誤って削除しない。
* ユーザーが明示していない仕様を推測だけで追加しない。
* 不明点が存在する場合でも、確定している範囲の調査・検討・作業は進める。
* 既存ドキュメントの文体、章構成、用語、項番体系を尊重する。
* 変更時は、上流工程と下流工程の両方向への影響を確認する。
* 修正後は、変更前情報の残存、項番、ID、参照関係、ドキュメント間の整合性を確認する。

# 回答方針

* まず結論を示し、その後に理由を説明する。
* 現在の案に問題がなければ、無理に変更案を作らず、その旨を明示する。
* 複数案がある場合は、推奨案を明確にする。
* 理論上の理想構成より、実際の作業効率、保守性、再利用性を重視する。
* 情報不足の場合は、確認できた範囲と確認できない範囲を区別する。

現時点では、この程度で十分です。

`doc-sync.agent.md` の全文をProject instructionsへ入れないことをおすすめします。

---

# 6. Projectにアップロードするファイル

最初から全ファイルを無条件にアップロードする必要はありません。

最初は、

```text
requirements.md
basic_design.md
function_spec.md
detail_design.md
test_design.md
test_spec.md
operation_design.md
operation_manual.md
```

など、実際に変更波及の対象になるドキュメントを追加します。

ChatGPT Projectには、PDF、スプレッドシート、文書、画像をアップロードしたり、テキストを貼り付けたりできます。追加したファイルはProject sourceとして利用されます。([OpenAI Help Center][1])

## 私なら最初はこうします

```text
Project Sources

requirements.md
basic_design.md
function_spec.md
detail_design.md
test_design.md
test_spec.md
operation_design.md
operation_manual.md
```

一方、

```text
src/
```

以下のソースコード全体は、最初からChatGPT Projectに持ち込まなくてもよいと思います。

コード全体の横断検索・直接編集はVS Code + Copilot側の方が自然だからです。

---

# 7. ChatGPT Project内のチャット構成

前回は、

```text
① 変更影響調査
② ドキュメント同期・修正
③ 構成改善レビュー
```

を提案しました。

少し具体化すると、私は次の4チャットから始めます。

```text
01_相談・方針検討

02_変更影響調査

03_ドキュメント同期・レビュー

04_ChatGPT運用改善
```

## 01_相談・方針検討

例えば、

```text
実装中にJSON出力からXML出力へ変更する必要が出ました。

理由は外部ライブラリ使用禁止の顧客制約です。

この変更を仕様変更として扱うべきか、
実装方式変更として扱うべきか検討してください。
また、どの工程のドキュメントへ影響する可能性があるか、
修正前に整理してください。
```

のような相談用です。

ここではまだ修正をさせません。

---

## 02_変更影響調査

こちらは影響調査専用です。

ローカルの、

```text
chatgpt/workflows/impact-analysis.md
```

に定型文を保存しておきます。

今回の変更について、プロジェクト内の関連ドキュメントへの影響を調査してください。

この段階ではドキュメントを修正せず、影響調査のみ行ってください。

以下を整理してください。

1. 変更内容の整理
2. 影響する可能性があるドキュメント
3. 影響する章・項目
4. 想定される修正内容
5. 確認したが修正不要と考えられる内容
6. 上流工程への影響
7. 下流工程への影響
8. 要確認事項
9. 残存リスク

変更前の用語が存在するという理由だけで修正対象とは判断せず、変更履歴、調査経緯、不採用理由、比較説明などとして残す必要がないかも確認してください。

これをローカルでGit管理します。

---

## 03_ドキュメント同期・レビュー

こちらは、影響調査後の具体的な修正検討用です。

例えば、

```text
先ほどの影響調査結果を基準に、
関連ドキュメントの修正内容を検討してください。

変更対象ごとに、
・現在の記載
・問題点
・修正案
・他ドキュメントへの影響
を整理してください。

最後に、
・変更前仕様の残存
・項番とID
・参照関係
・要件→設計→テストの整合性
・設計→運用の整合性
を再確認してください。
```

とします。

---

## 04_ChatGPT運用改善

これは、`copilot-config-reviewer.agent.md` のChatGPT版に近いです。

数回使ってから、

```text
このProjectの現在の運用構成をレビューしてください。

対象:
・Project instructions
・Project Sources
・チャットの分け方
・定型プロンプト
・ローカル側のchatgptフォルダ構成

以下の観点で確認してください。
・コンテキスト効率
・作業効率
・再利用性
・保守性
・役割分担

改善効果が明確でない場合は、
現状維持を推奨してください。
```

とレビューさせます。

---

# 8. 実際の1回の変更作業フロー

今回のJSON → XMLの例で、具体的に流すとこうなります。

## ローカル：変更要求を作る

```text
docs/
└── 00_change_requests/
    └── CR-0001_json-to-xml.md
```

内容：

```md
# CR-0001 ファイル出力形式変更

## 変更種別

仕様変更 / 実装制約対応

## 変更前

JSON形式でファイルを出力する。

## 変更後

XML形式でファイルを出力する。

## 変更理由

顧客制約により外部ライブラリを使用できない。
標準機能による実装と保守性を考慮し、XML形式へ変更する。

## 未確定事項

- XMLスキーマの有無
- 文字コード
- 改行コード
- ファイル名規則
```

---

## ChatGPT：影響調査

`02_変更影響調査` チャットで、

```text
CR-0001相当の変更を検討しています。

変更前:
JSON形式

変更後:
XML形式

理由:
顧客制約により外部ライブラリを使用できないため。

Project Sourcesのドキュメントを横断的に確認し、
影響調査のみ行ってください。
まだ修正案の確定やファイル修正は行わないでください。
```

と依頼します。

---

## ChatGPT：調査結果を確認

例えば結果が、

```text
基本設計書
→ ファイル連携方式の変更

機能仕様書
→ 出力形式、拡張子変更

詳細設計書
→ シリアライズ処理変更

テスト設計書
→ XML構造・文字コード観点追加

テスト仕様書
→ 期待ファイル変更

運用手順書
→ 確認対象拡張子変更
```

となったら、その結果を確認します。

有用な分析結果は、ChatGPTの回答メニューからProject sourceとして保存できます。公式にも、要約・意思決定メモ・下書き・分析結果などをProject sourceとして保存できる手順が案内されています。([OpenAI Help Center][1])

ただし、私は**毎回保存しすぎない**方がよいと思います。

重要な意思決定だけを保存します。

---

# 9. ChatGPTの結果をローカルに戻す方法

ここは重要です。

私は、

```text
ChatGPTで検討
    ↓
VS Codeで実際に修正
    ↓
Git diff
    ↓
必要ならChatGPTで再レビュー
```

をおすすめします。

つまり、

```text
ChatGPTが正本
```

にはしません。

ローカルを正本にします。

例えばChatGPTで確定した重要な判断は、

```text
chatgpt/logs/decisions.md
```

または、

```text
docs/00_change_requests/
CR-0001_json-to-xml.md
```

に戻します。

---

# 10. `decisions.md` はかなり有効だと思います

例えば、

```text
chatgpt/
└── logs/
    └── decisions.md
```

を作ります。

内容は、

```md
# Decision Log

## 2026-07-05 CR-0001

### 決定

ファイル出力形式をJSONからXMLへ変更する。

### 理由

顧客制約により外部ライブラリを利用できないため。

### 影響範囲

- 基本設計
- 機能仕様
- 詳細設計
- テスト設計
- テスト仕様
- 運用手順

### 未確定

- XML Schemaの作成有無
- 文字コード
```

という形です。

これをGit管理しておくと、後で、

> なぜこの判断をしたのか

が分かりやすくなります。

---

# 11. ファイル更新時のおすすめ運用

ChatGPT Projectにはローカルフォルダとの自動同期を前提にした機能ではなく、Project sourceとしてファイルを追加して使う仕組みです。Projectへ追加済みのファイルと同名のファイルを再アップロードする際には、そのまま追加するかスキップするかを選択する挙動が公式ヘルプに記載されています。([OpenAI Help Center][1])

したがって、ローカル側のドキュメントを頻繁に更新する場合は、**Project sourceの更新管理を意識する必要があります**。

私は最初は、

```text
大きな変更前
↓
必要なドキュメントをProjectへ追加

検討・レビュー
↓
ローカル修正

一連の変更完了
↓
必要に応じてProjectの資料も更新
```

くらいで始めることをおすすめします。

---

# 私の推奨する最終構成

## ローカル

```text
software-project/
│
├── .github/
│   ├── agents/
│   │   ├── doc-sync.agent.md
│   │   └── copilot-config-reviewer.agent.md
│   └── copilot-instructions.md
│
├── docs/
│   ├── 00_change_requests/
│   ├── 01_research/
│   ├── 02_estimate/
│   ├── 03_requirements/
│   ├── 04_basic_design/
│   ├── 05_function_spec/
│   ├── 06_detail_design/
│   ├── 07_test_design/
│   ├── 08_test_spec/
│   ├── 09_operation_design/
│   └── 10_operation_manual/
│
├── chatgpt/
│   ├── project-instructions.md
│   ├── workflows/
│   │   ├── impact-analysis.md
│   │   ├── doc-sync.md
│   │   └── config-review.md
│   └── logs/
│       └── decisions.md
│
├── src/
├── tests/
└── README.md
```

## ChatGPT Project

```text
Project: 開発ドキュメント管理

├─ Project Instructions
│
├─ Sources
│   ├─ 要件定義書
│   ├─ 基本設計書
│   ├─ 機能仕様書
│   ├─ 詳細設計書
│   ├─ テスト設計書
│   ├─ テスト仕様書
│   ├─ 運用設計書
│   └─ 運用手順書
│
├─ Chat: 01_相談・方針検討
├─ Chat: 02_変更影響調査
├─ Chat: 03_ドキュメント同期・レビュー
└─ Chat: 04_ChatGPT運用改善
```

私の判断では、**この構成から始めるのがちょうどよいです**。

ToshiさんはこれまでCopilotについても「まず小さく始め、実際に使ってからAgent・Skill・Promptへ分割する」という進め方を検討してきましたが、ChatGPT側も同じ方針が合っています。最初からCustom GPTまで作らず、まずProject＋4チャット＋ローカル`chatgpt/`管理で運用し、繰り返しパターンが見えてからCustom GPT化を判断するのがよいと思います。

[1]: https://help.openai.com/ja-jp/articles/10169521-projects-in-chatgpt "ChatGPT のプロジェクト | OpenAI Help Center"
[2]: https://help.openai.com/en/articles/10169521-projects-in-chatgpt "Projects in ChatGPT | OpenAI Help Center"

