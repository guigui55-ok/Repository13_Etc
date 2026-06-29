# Goal
Github Copilot の SKILL.md の学習（お試し動作確認をする）

# Details
- copilot-instructions.md と SKILL.md のみの最小構成とする。

# 題材
題材は 文章を「作業メモ形式」に整えるスキル

# 構成例：

your-workspace/
├─ .github/
│  └─ copilot-instructions.md
└─ .github/
   └─ skills/
      └─ simple-work-memo/
         └─ SKILL.md

# 使い方
VSCodeでのテスト

Copilot Chat にこう入れます。

workフォルダのメモを整理してください。
output.mdに反映してください。

期待挙動：

[SKILL:memo-reader]

変更したファイル：
- work/output.md

そして work/output.md が更新される。

↓
実際の挙動
ファイルは作成されたが、スキルは使用されない。


# 使い方2
work/memo.txt を読んで整理し、work/output.md に反映してください。

この作業では SKILL.md を使ってください。
↓
作業結果
output_sample1.md
↓
ファイル作成・スキル使用OK
スキル使用を明示すれば使用される。

# 使い方3
work/memo.txt を整理して output.md に反映



# メモ
[SKILL.md](file:///c%3A/Users/guigu/Desktop/260507%20copilot/sample_skills/.github/skills/memo-reader/SKILL.md) にトリガー例と出力サンプルを追加


