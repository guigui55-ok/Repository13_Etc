可能です。
ただし、**「通常利用の誤接続防止」としては有効**ですが、**セキュリティ的に強い認証ではありません**。

結論としては、方針はだいたい合っています。

## 判定

### 既存クライアント → 別Editionサーバー

別Edition側ストアドで、

```sql
@EditionId varchar(50)
```

のような必須引数を追加し、既存クライアントがその引数を渡さないなら、

* ストアド呼び出しエラー
* または認証NG

にできます。

これは可能です。

### 別Editionクライアント → 既存サーバー

別Editionクライアントが新しい引数を渡そうとすると、既存ストアドにはその引数がないため、

* “too many arguments specified”
* 戻り値にEdition識別子がない
* 期待する戻り値形式と違う

などで、別Editionクライアント側で「既存サーバーなのでNG」と判定できます。

これも可能です。

## 注意点

一番重要なのは、**識別子をクライアントから渡す方式は偽装可能**という点です。
つまり、別Edition識別子がクライアント内に入っているなら、解析・改造されれば既存クライアント側からも同じ値を渡せます。

なので、この方式は、

> 間違って別Editionサーバーにつなぐことを防ぐ

用途ならOKです。

一方で、

> 悪意ある利用者や改造クライアントも確実に防ぐ

用途には弱いです。

## 改善案

ストアド内定数でも動きますが、個人的には以下の方が管理しやすいです。

```text
EditionSettings テーブル
- EditionId
- ServerEditionName
- ProtocolVersion
```

ログインストアドでは、

1. クライアントから `ClientEditionId` を受け取る
2. DB側の `EditionId` と比較
3. 不一致なら認証NG
4. 一致ならログイン成功
5. 戻り値に `ServerEditionId` / `ProtocolVersion` を返す

という形です。

## さらに重要な点

今回、ログイン前に

> DBのテーブル一覧を取得している

とのことなので、**ログインストアドだけで制限しても、テーブル一覧取得までは防げません**。

既存クライアントが別Edition DBに接続できるSQLユーザーを持っているなら、ログイン前のDBアクセス自体は発生します。

本当に防ぎたいなら、別Edition DB側で、

* 既存クライアント用SQLログインを作らない
* 接続文字列を別にする
* SQL権限をストアド実行だけに絞る
* テーブル一覧取得権限を与えない

あたりも検討した方がよいです。

## まとめ

提示された方式は**可能**です。
ただし位置づけは、

```text
Edition不一致の誤接続防止：OK
強固なセキュリティ対策：不十分
```

です。

実装するなら、最低限、

```text
別Editionログインストアドに必須のEditionId引数を追加
DB側EditionIdと比較
戻り値にもServerEditionIdを返す
別Editionクライアント側でも戻り値を検証
```

でよいと思います。


User:
既存クライアント → 別Editionサーバー　について、引数違いによるストアド呼び出しエラー、になるか、認証NGになるかは動かしてみないとわかりませんかね？ 別Editionクライアント → 既存サーバー　についても、回答くださった中のどの挙動になるか、今はわからないですよね？ 悪意ある利用者や改造クライアントも確実に防ぎたいところですが、実装とテストの工数があまりなく、上記の実装としたいです。


ChatGpt:
はい、その理解でよいです。

## 既存クライアント → 別Editionサーバー

これは**実装次第**です。

別Editionストアドをこう変更した場合：

```sql
@EditionId varchar(50)
```

かつデフォルト値なしなら、既存クライアントは引数を渡せないため、基本的には**ストアド呼び出しエラー**になります。

```sql
@EditionId varchar(50) = NULL
```

のようにデフォルト値ありにすると、呼び出し自体は成功し、ストアド内で、

```sql
IF @EditionId IS NULL
    RETURN 認証NG
```

のようにできます。

つまり、

```text
引数デフォルトなし → 呼び出しエラー寄り
引数デフォルトあり → 認証NGに制御しやすい
```

です。

## 別Editionクライアント → 既存サーバー

こちらは今の既存ストアドの呼び出し方法次第です。

別Editionクライアントが新しい引数を渡すなら、既存ストアド側にはその引数がないので、通常は

```text
引数が多すぎる
指定された引数が存在しない
```

系のエラーになります。

ただし、戻り値の受け取り方や呼び出しラッパーによって、別の例外・失敗扱いになる可能性はあります。

なので、こちらは**実機または検証DBで確認した方が安全**です。

## 工数が少ない場合のおすすめ

今回の工数制約なら、方針はこれでよいと思います。

```text
1. 別Editionストアドには @EditionId = NULL を追加
2. NULL または不一致なら認証NG
3. 一致なら認証OK
4. 戻り値に ServerEditionId を追加
5. 別Editionクライアント側で ServerEditionId を確認
6. 既存サーバー接続時の例外は「Edition不一致」として扱う
```

ポイントは、**ストアド呼び出しエラーを正常な判定材料にしすぎない**ことです。

別Editionサーバー側は、できればエラーではなく、

```text
認証NG
Edition不一致
```

として返せるようにした方が、テストもしやすいです。

結論としては、
**強固なセキュリティではないが、誤接続防止としては現実的な実装**です。


User:
以下の形が一番扱いやすいです。

## 変更方針

既存ストアドを別Edition側でだけ変更します。

```sql
@ClientEditionId varchar(50) = NULL
```

のように、**デフォルト NULL 付きで引数追加**します。

これにより、

* 既存クライアント → 引数を渡さない → `NULL` → 認証NG
* 別Editionクライアント → 正しい識別子を渡す → 認証処理へ進む

にできます。

## ストアド例

例として、ログインストアドがこういうものだとします。

```sql
CREATE PROCEDURE dbo.LoginUser
    @UserId varchar(50),
    @Password varchar(100)
AS
BEGIN
    SET NOCOUNT ON;

    -- ユーザー認証処理
    SELECT
        UserId,
        UserName,
        1 AS LoginResult
    FROM dbo.Users
    WHERE UserId = @UserId
      AND Password = @Password;
END
```

別Edition側では、次のように変更します。

```sql
ALTER PROCEDURE dbo.LoginUser
    @UserId varchar(50),
    @Password varchar(100),
    @ClientEditionId varchar(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @ServerEditionId varchar(50) = '別Edition識別子';

    -- Edition不一致
    IF @ClientEditionId IS NULL
       OR @ClientEditionId <> @ServerEditionId
    BEGIN
        SELECT
            0 AS LoginResult,
            'EDITION_MISMATCH' AS ErrorCode,
            @ServerEditionId AS ServerEditionId;

        RETURN;
    END

    -- 通常のユーザー認証
    IF EXISTS (
        SELECT 1
        FROM dbo.Users
        WHERE UserId = @UserId
          AND Password = @Password
    )
    BEGIN
        SELECT
            1 AS LoginResult,
            NULL AS ErrorCode,
            @ServerEditionId AS ServerEditionId,
            @UserId AS UserId;
    END
    ELSE
    BEGIN
        SELECT
            0 AS LoginResult,
            'LOGIN_FAILED' AS ErrorCode,
            @ServerEditionId AS ServerEditionId;
    END
END
```

## ポイント

`@ClientEditionId varchar(50) = NULL` にするのが重要です。

デフォルト値なしにすると、既存クライアントから呼ばれた時点でストアド呼び出しエラーになりやすいです。

```sql
@ClientEditionId varchar(50)
```

よりも、

```sql
@ClientEditionId varchar(50) = NULL
```

の方が、**認証NGとして制御しやすい**です。

## 戻り値の考え方

最低限、以下を返すとよいです。

```text
LoginResult
ErrorCode
ServerEditionId
```

例：

```sql
SELECT
    0 AS LoginResult,
    'EDITION_MISMATCH' AS ErrorCode,
    @ServerEditionId AS ServerEditionId;
```

別Editionクライアント側では、

```text
LoginResult = 1
かつ
ServerEditionId = 期待する別Edition識別子
```

の場合だけログイン成功にします。

## 識別子の例

```sql
DECLARE @ServerEditionId varchar(50) = 'EDITION_B_2026';
```

この程度でよいです。

ただし、簡単に推測されにくくしたいなら、

```sql
DECLARE @ServerEditionId varchar(100) = 'EDITION_B_2026_xxxxxxxx';
```

のように少し長めにしてもよいです。

## 実装順

```text
1. 別Editionストアドに @ClientEditionId = NULL を追加
2. ストアド先頭で識別子チェック
3. 不一致なら LoginResult=0 を返す
4. 一致した場合だけ既存の認証処理へ進む
5. 戻り値に ServerEditionId を追加
6. 別Editionクライアントで ServerEditionId を確認
```

この形なら、工数少なめで実装しやすいです。


User:
ストアドの登録方法を、sqlを用いて行いたいです。

Chat:
SQL Serverなら、基本は **`ALTER PROCEDURE`** で既存ストアドを更新します。

```sql
ALTER PROCEDURE dbo.LoginUser
    @UserId varchar(50),
    @Password varchar(100),
    @ClientEditionId varchar(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @ServerEditionId varchar(50) = 'EDITION_B_2026';

    IF @ClientEditionId IS NULL
       OR @ClientEditionId <> @ServerEditionId
    BEGIN
        SELECT
            0 AS LoginResult,
            'EDITION_MISMATCH' AS ErrorCode,
            @ServerEditionId AS ServerEditionId;
        RETURN;
    END

    -- ここから既存のログイン認証処理
    SELECT
        1 AS LoginResult,
        NULL AS ErrorCode,
        @ServerEditionId AS ServerEditionId;
END
GO
```

新規作成も兼ねたいなら、SQL Server 2016 SP1以降ではこれが使えます。

```sql
CREATE OR ALTER PROCEDURE dbo.LoginUser
    @UserId varchar(50),
    @Password varchar(100),
    @ClientEditionId varchar(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @ServerEditionId varchar(50) = 'EDITION_B_2026';

    IF @ClientEditionId IS NULL
       OR @ClientEditionId <> @ServerEditionId
    BEGIN
        SELECT
            0 AS LoginResult,
            'EDITION_MISMATCH' AS ErrorCode,
            @ServerEditionId AS ServerEditionId;
        RETURN;
    END

    -- 既存認証処理
    SELECT
        1 AS LoginResult,
        NULL AS ErrorCode,
        @ServerEditionId AS ServerEditionId;
END
GO
```

既存ストアドを安全に置き換えるなら、SSMSで以下の流れです。

```sql
USE 対象DB名;
GO

CREATE OR ALTER PROCEDURE dbo.ストアド名
    -- 既存引数,
    @ClientEditionId varchar(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    -- 変更後の処理
END
GO
```

確認用SQLはこれです。

```sql
EXEC dbo.LoginUser
    @UserId = 'test',
    @Password = 'pass',
    @ClientEditionId = 'EDITION_B_2026';
```

既存クライアント相当の確認はこれです。

```sql
EXEC dbo.LoginUser
    @UserId = 'test',
    @Password = 'pass';
```

この場合、`EDITION_MISMATCH` が返れば想定通りです。

