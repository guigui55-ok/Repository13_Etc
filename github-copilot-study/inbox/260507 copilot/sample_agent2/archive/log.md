
----
User: work/memo.txt を整理してください
↓
Agent,skill使用せず

----
memo.txt を整理して output.md に反映してください。

AGENT.md に従って memo.txt を整理してください。
結果は output.md に保存してください。
↓
agentのみ使用

-----
memo.txt を整理して output.md に反映してください。
カスタムエージェント選択
↓
agent使用していない？
↓
skillのformatterの修正でスキルの発動ができた


-----

このファイルを整理して、ファイル出力してください（memo-organizer.agent.md, 関連スキルを使用）


-------

※注意点

SKILL.md の frontmatter が壊れていると、発動しない
--- で始めたら、閉じも --- です。
→ハイフン3つを使うのが安全

作業が簡単すぎたりすると、Skillなしでも解けてしまう可能性あり。
他に、似たようなskillがあったり、作業依頼があいまいだと、思い通りに発動しなさそう。

