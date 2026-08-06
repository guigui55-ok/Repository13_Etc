# 目的
C# asp.net のサーバーアプリとクライアントアプリの連携の動作確認をする。
特に、サーバー側WebAPI（asmx Web サービス）と クライアント側の呼び出し部分をInvokeで実装し、この仕様を理解する。
クライアント側で引数を合わせなかったときにSOAPエラーとなることも確認したい。

* 成功条件
TestAspClientDefault 実行後、サーバー側の各 WebMethod が正常に呼ばれ、コンソールに期待される結果が出力されること。


# 詳細
サーバー側はWebAPIを公開し、クライアント側とデータのやり取りを行う。
いくつかの引数のパターンを確認する。

サーバー・クライアントはローカルPCで完結させる。

# 構成など
* サーバー側
プロジェクト名: TestAspServerDefault  
言語など: C#, .Net Framework 4.7, ASP.Net  
プロジェクト種別: ASP.NET Web アプリケーション (.NET Frameowk)
テンプレート: Web Forms テンプレート — ASMX（SOAP）Web サービスをすぐに追加・動作確認できるため。

* クライアント側
プロジェクト名: TestAspClientDefault  
言語など: C#, .Net Framework 4.7  
アプリ形態: コンソールアプリ  

* コンポーネント
Visual Studio Installer コンポーネント追加
    個別コンポーネント
	    .NET Framework 4.6.2-4.7.1 開発ツール
    Asp.Net インストールの詳細
        .Net Framework プロジェクトと項目テンプレート

* 通信方法
    SOAP（asmx）

* ※ 備考
    Defaultと名付けた理由は、今後Aspの別のTest系プロジェクトが作る可能性があるた
    め。  
    職場での環境に合わせ C# .Net Framework 4.7 を採用  

* フォルダ・ファイル構成（ツリー）  
本ワークスペースのroot  
    LTestAspServerDefault  
        LTestAspServerDefaultプロジェクトファイル  
        Lその他ソースファイルなど  
    LTestAspClientDefault  
        LTestAspClientDefault.csprj  
        Lその他ソースファイルなど  
※ツリーの詳細は確認中のため概略の未記載。

# フロー
クライアント側で  
TestAspClientDefault実行  
    実行後すぐにWebAPIを呼び出し  
↓  
サーバー側で  
TestAspClientDefault が呼び出されたWebAPI処理が実行される  
↓  
クライアント側で  
結果や情報を受け取り、コンソールに出力  
↓  
終了  

# WebAPI
* 本学習はサーバー引数の指定方法、Invokeの仕様を学習したいため以下のような構成にする。
* 基本的にAPIの戻り値は結果クラスResultInfo（後述を参考）をreturnし、その他の情報は、out引数でクライアントに渡す。

以下、確認したいパターンを列挙する。
1. 引数無し
    Webメソッド: TestWebMethod_Default()
2. 引数1つ out string
    Webメソッド: TestWebMethod_OutOne()
3. 引数1つ string
    Webメソッド: TestWebMethod_InOne()
4. 引数2つ out string, string
    Webメソッド: TestWebMethod_OutIn()
5. 引数2つ string, string
    Webメソッド: TestWebMethod_InIn()
6. 引数3つ out string, out string, string
    Webメソッド: TestWebMethod_OutOutIn()
7. 引数3つ string, out string,out string
    Webメソッド: TestWebMethod_InOutOut()
8. 引数3つ string, out string, string
    Webメソッド: TestWebMethod_InOutIn()
9. 引数3つ out string, string, out string
    Webメソッド: TestWebMethod_OutInOut()

# クライアント呼び出し
Invokeを使用してWebAPIを呼び出す。

# 結果コードについて
独自クラス ResultInfo を使用する。
サーバー側、クライアント側で、同じ解釈となるように実装する。
（片方でOKを返し、片方でNGとなったりしないこと）

* ResultInfo メンバ
    ResultCode Result
    string Data
※とりあえず1つのみ

* 結果コードを扱うEnum
    Enum ResultCode
        None
        OK
        Fail
        Error
        Unexpected

# ASMX Web サービスの Invoke 呼び出しサンプル
* HttpSimpleClientProtocol.Invoke を使用する。
https://learn.microsoft.com/ja-jp/dotnet/api/system.web.services.protocols.httpsimpleclientprotocol.invoke?view=netframework-4.8.1


# サーバー側 Asp.net 実装
* 空テンプレート
1. Service.asmx と Service.asmx.cs
    Service.asmx は ASMX Web サービスのエントリ
    Service.asmx.cs に WebMethod を実装
    ※ソリューション、追加＞asmx（単純ファイル追加だと、csprojも手動追加が必要）
2. 共有クラス（契約）を置く場所
    App_Code フォルダを作成して、ここに ResultInfo.cs と ResultCode.cs を置くのが分かりやすい
    またはプロジェクト直下の Models / Contracts フォルダでも可
    ※ソリューションエクスプローラー、右クリック「Asp.net フォルダの追加」で追加
3. Service.asmx.cs に 9 種類のメソッドを実装
* その他
ResultInfoの参照エラーが発生するときは、csprojの<Compile>を確認
    <Compile Include="App_Code\ResultInfo.cs" />
4. サーバーSSL設定無効
    csproj
        URLをhttps→httpに
    applicationhost.config
    <bindings>
        <binding protocol="http" bindingInformation="*:51582:localhost" />
    </bindings>

5. サーバー起動確認
    F5（or デバッグ実行）→ブラウザでエラーが表示されないことを確認

    ビルド時の以下のダイアログは「いいえ」を選択。
``` ビルド時のDialog
このプロジェクトはSSLを使用するように構成されています。ブラウザーでのSSLの警告を避けるには、IIS Expressが生成した自己署名証明書を信頼することを選択します。
IIS Express の SSL 証明書を信頼しますか？
[はい] [いいえ]
```

6. サーバーデプロイ

7. クライアント実装
参照追加
    System.Web.Services
コード実装
    割愛
参考
https://qiita.com/junkichi424/items/2a4fad9f0bdeeee74e1d

# 残タスク
各メソッドのFail版作成
引数を異なることによるSoapエラーを発生させる（不可能かもしれない）
引数情報によって、呼び出し元を判定し、条件によってエラーにする
　（WebAPIバージョン（IF）が違う）