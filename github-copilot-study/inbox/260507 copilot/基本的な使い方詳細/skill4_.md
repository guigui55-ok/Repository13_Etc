# skillについて調査

## 公式
https://docs.github.com/ja/copilot/concepts/agents/about-agent-skills


## any memo

### 使用手順

* プロジェクト スキルの場合
リポジトリに格納されているProjectスキル (.github/skills、.claude/skills、または .agents/skills)

* 個人のスキルの場合
ホーム ディレクトリに格納され、プロジェクト間で共有される個人スキル (~/.copilot/skills または ~/.agents/skills)

* ホームディレクトリ
Windows  
    C:\Users\<ユーザー名>  

* skills配置先
C:\Users\<ユーザー名>\.copilot\skills\ 
または C:\Users\<ユーザー名>\.agents\skills\

* 新しいスキルのサブディレクトリを作成
     (たとえば、 .github/skills/webapp-testing) 
     スキルサブディレクトリ名は小文字で、スペースにはハイフンを使用する必要があります。
     [重要]スキル ファイルには、 SKILL.mdという名前を付ける必要があります。

*
SKILL.md ファイルは、YAML frontmatter を含む Markdown ファイルです。 
最も単純な形式では、次のものが含まれます。

YAMLフロントマター
name (必須): スキルのユニークな識別子。 スペースにはハイフンを使用して、小文字にする必要があります。 通常、これはスキルのディレクトリの名前と一致します。
description (必須): スキルが実行する内容と、それを使用する必要があるタイミング Copilot 説明。
license (省略可能): このスキルに適用されるライセンスの説明。