
例えば、あるc#プロジェクト（ソースコードと仕様書）があるとして、これの仕様調査する場合

# 構成目安
| 種類         |   粒度 | 役割                                          |
| ---------- | ---: | ------------------------------------------- |
| **agent**  |  大きい | 「調査担当」「実装修正担当」「レビュー担当」など、役割・判断方針            |
| **skill**  | 中くらい | 「C#コード調査手順」「影響範囲調査手順」「仕様書照合手順」など、再利用できる作業手順 |
| **prompt** |  小さい | 今回だけの具体的な依頼                                 |


GitHub の公式ドキュメント上でも、Skill は SKILL.md を持つサブディレクトリとして作り、手順や補助ファイルをまとめるもの、という扱いです。 VS Code 側でも Agent Skills は複数の AI エージェントで使える標準として説明されています。


## 使い分けの結論

このケースでは、

**agent は 1個**

* `code-investigation-agent`

**skill は調査パターンごとに複数**

* `csharp-file-write-investigation`
* `csharp-version-check-investigation`
* `csharp-spec-cross-reference`
* `csharp-impact-analysis`

のようにするのがよいです。

依頼時は例えば、

```text
code-investigation-agent と csharp-file-write-investigation skill を使って、
ファイル出力処理に暗号化を追加する前提で、書き込み処理を調査し列挙してください。
結果は work/file-write-investigation.md に出力してください。
```

という感じです。

考え方としては、**agent は担当者、skill は作業手順書**です。
なので「仕様調査担当エージェント」が、「ファイル書き込み調査スキル」や「バージョン判定調査スキル」を使う、という構成が一番わかりやすいです。


agentについて
    環境や状況を伝える内容は、調査用、コーディング用に別に用意した方がよい
