* copilot-instructions.md（このプロジェクトではどう振る舞うべきか）
* *.instructions.md（特定ファイル群にだけ適用。）
* AGENTS.md（AIエージェントの行動方針）
* Custom Agents（*.agent.md）（特定専門AIを作る）
* SKILL.md（Agent Skills）（特定タスク知識をモジュール化」）
* Prompt Files（*.prompt.md）（再利用可能プロンプト」）
* MCP（AIへ外部ツールを接続）
* Hooks（AI処理前後にスクリプト実行）

* Copilot Coding Agent
* Spaces
* awesome-copilot

ChatGPTのツールも使用したい（ブラウザ版）


前提：
今仕事でGithubCopilotを使用しています。（プレミア枠あり）

私は、Copilotには主に以下の機能があると思っています。
* copilot-instructions.md（このプロジェクトではどう振る舞うべきか）
* *.instructions.md（特定ファイル群にだけ適用。）
* AGENTS.md
* Custom Agents（*.agent.md）（特定専門AIを作る）
* SKILL.md（Agent Skills）（特定タスク知識をモジュール化」）
* Prompt Files（*.prompt.md）（再利用可能プロンプト」）
* MCP（AIへ外部ツールを接続）
* Hooks（AI処理前後にスクリプト実行）
* Copilot Workspace
* GithubCLI

仕事では既存アプリ（大規模）があり、機能追加時の仕様の調査や設計検討、実装に使用しています。
機能追加時の仕様の調査と機能追加の設計をするときに、上記の機能をいくつか使えばかなり効率的に作業できると思っています。
どのようにしたらよいか、提案していただけますか？


----------
プライベートではフリーで使用しています。
仕事で効率的に作業をするために、Copilotの勉強をしたいです。


プライベートでも小規模なアプリを作り、体感しながらの学習をしようと思っています。

--------------

instructions.mdは毎回読み込むんですよね？
コンテキストが多くなりそうに思うんですが、そんなことはないんですかね？（最近はこれくらいが普通ですか？）

あと、調査について、普通の調査、簡易的な調査（呼び出し元だけを知りたい、など）、詳細な調査、など条件を指定したい場合があります、この時はどうしたらよいですか？

このようなmdファイルがいろいろありますが、質問時にどのように指定しますか？（質問内容を勝手に解釈して適切なものが使用されますか？）


--------------
# Instructions.md
C:\Users\guigu\Desktop\260507 copilot\example1\調査\.github\instructions\csharp.instructions.md

.instructions.md では、applyTo は実質ほぼ必須
applyTo に一致したファイルをCopilotが扱う時、その instructions が「候補として適用される」というイメージ
ただし、「絶対100%毎回読み込まれる」保証はない。
（内部的には relevance（関連度） / context window（AIが一度に読める情報量の上限） / task relevance（タスク関連度：今回の依頼内容に本当に必要か） で選別される）

まず .instructions.md の役割
これは、「特定ファイル群向けの追加ルール」です。
applyTo が無いと？　→　挙動が曖昧になる
    結果：効かなかったり、relevance低下したり、適用優先度が落ちたり、しやすいです。
applyTo は glob
    WinFormsだけ
    applyTo: "**/*Form*.cs"
    DB関連だけ
    applyTo: "DB/**/*.sql"
    複数
    applyTo:
    - "ClientApp/**/*.cs"
    - "SharedLib/**/*.cs"

ただし「強すぎる applyTo」は危険
例えば：

applyTo: "**/*"

に大量instructionsを書くと、

全部の作業で毎回参照候補になり、

コンテキスト圧迫
relevance低下
ノイズ化

しやすいです。

--------------
# prompt.md
作業ごとに分けるとよい。（要件・機能ごと、など）
質問時に /prompt/hoge.prompt.md のように指定すると読み込まれる。

prompt.md は、基本的に「自動では読み込まれない」（＝明示的に使う前提）


## その他
フォルダ分け可能
--------------

--------------
これまで、使用してきた質問の統計を取って、必要なものをmdに落とし込む
    →抒情に自動化


| S   | copilot-instructions.md |
| S   | Agent Mode              | →VsCode質問欄のAsk,Plan,AgentのAgentのこと
| A   | SKILL.md                |
| A   | AGENTS.md               |
| A   | Prompt Files            |
| B   | MCP                     | 260517 発展途上
| B   | Hooks                   | 260517 発展途上
| B   | Custom Agents           |

# 学習予定

instructions.md
AGENTS.md
skills.md
prompt.md

AgentModeのAgentの使用方法

## 保留
MCP

prompt,skills,agent,instructions 4種類のmdがありますが、これは、状況によって複数の種類のmdが読み込まれることがある、同じ種類でも複数のmdが読み込まれる可能性がある認識で合っていますか？





