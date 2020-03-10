---
languages:
- javascript
page_type: sample
description: "このサンプルは、Outlook リボンにボタンを追加する Outlook アドインを実装します。"
products:
- office
- office-outlook
urlFragment: outlook-ribbon-addin
---

# AttachmentsDemo サンプル Outlook アドイン

このサンプルは、Outlook リボンにボタンを追加する Outlook アドインを実装します。このボタンを使用すると、ユーザーはすべての添付ファイルを OneDrive に保存できるようになります。このサンプルでは次の概念を示します。
 
- メールを表示するときに、UI を使用しないボタンおよび作業ウィンドウを開くボタンを含む、[アドイン コマンド ボタン](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)を Outlook に追加する。
- [コールバック トークンおよび Outlook REST API を使用して添付ファイルを取得する](https://dev.office.com/docs/add-ins/outlook/use-rest-api)ための Web API を実装する。
- [SSO アクセス トークンを使用](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)して、ユーザーへの確認を表示せずに Microsoft Graph API を呼び出す。
- SSO トークンがない場合に、OAuth2 の暗黙のフローを [office-js-helpers ライブラリ](https://github.com/OfficeDev/office-js-helpers)経由で使用してユーザーの OneDrive への認証を行う。
- [Microsoft Graph API](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) を使用して OneDrive にファイルを作成する。

## サンプルの構成

このサンプルを正常に動作させるには、サンプルを実行する前にいくつかの操作を行う必要があります。

1. Office 365 テナントまたは Outlook.com アカウントが必要です。メール アプリは Exchange のオンプレミス インストールで動作しますが、Microsoft Graph API では Office 365 または Outlook.com が必要です。
2. Microsoft Graph API にアクセスするためのアプリ ID を取得するには、サンプル アプリケーションを [Microsoft アプリケーション登録ポータル](https://apps.dev.microsoft.com)で登録する必要があります。
    1. [Microsoft アプリケーション登録ポータル](https://apps.dev.microsoft.com)に移動します。サインインを求められない場合は、\[**アプリの一覧に移動**] ボタンをクリックし、Microsoft アカウント (Outlook.com) か職場または学校アカウント (Office 365) のいずれかを使用してサインインします。サインインしたら、\[**アプリの追加**] をクリックします。名前として「`AttachmentsDemo`」と入力し、\[**アプリケーションの作成**] をクリックします。
    1. \[**アプリケーション シークレット**] セクションを見つけ、\[**新しいパスワードを生成**] をクリックします。ダイアログ ボックスが開き、生成されたパスワードが表示されます。ダイアログ ボックスを閉じる前に、この値をコピーして保存します。
    1. \[**プラットフォーム**] セクションを見つけ、\[**プラットフォームの追加**] をクリックします。\[**リダイレクト URI**] で \[**Web**] を選択し、"`https://localhost:44349/MessageRead.html`" と入力します。
        > **注:**お客様の開発用コンピューターでは、リダイレクト URI のポート番号 (`44349`) が異なる場合があります。コンピューターの正しいポート番号を見つけるには、**ソリューション エクスプローラー**で \[**AttachmentDemoWeb**] プロジェクトを選択し、\[プロパティ] ウィンドウの \[**開発サーバー**] で \[**SSL URL**] 設定を確認します。
        
    1. \[**プラットフォームの追加**] をクリックします。\[**Web API**] を選択します。このセクションを次のように構成します。
        - \[**アプリケーション ID URI**] で、そこに入力されている GUID の前にホストとポート番号を挿入して既定値を変更します。たとえば、既定値が "`api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`" である場合、アプリが `localhost:44349` で実行されている場合の値は "`api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`" となります。
        - \[**事前承認済みアプリケーション**] で、\[**アプリケーション ID**] に "`d3590ed6-52b3-4102-aeff-aad2292ab01c`" と入力します。\[**スコープ**] ドロップダウンをクリックし、そこに 1 つだけ表示されているエンティティを選択します。これにより、(Windows の) デスクトップ版 Office がこのアプリにアクセスすることが事前承認されます。
        - \[**事前承認済みアプリケーション**] で、\[**アプリケーション ID**] に "`bc59ab01-8403-45c6-8796-ac3ef710b3e3`" と入力します。\[**スコープ**] ドロップダウンをクリックし、そこに 1 つだけ表示されているエンティティを選択します。これにより、Outlook on the web がこのアプリにアクセスすることが事前承認されます。
    1. アプリ登録で、\[**Microsoft Graph のアクセス許可**] セクションを見つけます。\[**委任されたアクセス許可**] の横にある \[**追加**] をクリックします。\[**Files.ReadWrite**]、\[**Mail.Read**]、\[**offline\_access**]、\[**openid**]、および \[**profile**] を選択します。\[**OK**] をクリックします。

\[**保存**] をクリックし、登録を完了します。\[**アプリケーション ID**] をコピーし、先ほど保存したアプリのパスワードと同じ場所に保存します。これらの値は、後で必要になります。

完了後のアプリケーション登録の詳細は、次のように表示されます。

![完了したアプリ登録](readme-images/app-registration.PNG)
![完了したアプリ登録その 2](readme-images/web-api-app-registration.PNG)

[authconfig.js](AttachmentDemoWeb/Scripts/authconfig.js) を編集し、\[`アプリ ID をここに入力してください`] の値をアプリケーション登録プロセスで生成したアプリケーション ID で置き換えます。

[AttachmentDemo.xml](AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml) を編集し、\[`アプリ ID をここに入力してください`] の値をアプリケーション登録プロセスで生成したアプリケーション ID で置き換えます。

> **注**: \[`リソース`] 要素のポート番号がプロジェクトで使用されているポートと一致していることを確認します。アプリケーションを登録するときに使用したポートとも一致する必要があります。

[Web.config](AttachmentDemoWeb/Web.config) を編集し、\[`アプリ ID をここに入力してください`] の値をアプリケーション ID で置き換え、\[`アプリ パスワードをここに入力してください`] をアプリケーション登録プロセスで生成したアプリケーション パスワードで置き換えます。

## ユーザーの同意をアプリに付与する

この手順では、先ほどアプリで構成したアクセス許可にユーザーの同意を付与します。この手順は、開発およびテスト用にアドインをサイドローディングするという理由から**のみ**必要です。通常、運用アドインは Office Store に掲載され、ユーザーは Store 経由のインストール プロセス中に同意を求められます。

同意するための選択肢は 2 つあります。Office 365 組織内のすべてのユーザーに一度に管理者のアカウントと同意を使用することができます。または、特定のユーザーの同意のために任意のアカウントを使用することができます。

### すべてのユーザーに管理者の同意を提供する

テナント管理者のアカウントにアクセスできる場合は、この方法を使用すると、組織内のすべてのユーザーに同意を提供できるようになります。アドインの開発とテストを複数の開発者で進める必要がある場合は、この方法が便利です。

1. `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345` に移動します。`{application_ID}` は、アプリケーション登録で表示されたアプリケーション ID です。
1. 管理者アカウントでサインインします。アクセス許可を確認してから ［**承諾**］ をクリックします。

ブラウザーは元のアプリへのリダイレクトを試行しますが、そのアプリは実行されていない可能性があります。［**承諾**］ をクリックした後に、「このサイトにアクセスできません」というエラーが表示されることがあります。これは問題ありません。この場合でも同意は記録されています。

### 単一のユーザーに同意を提供する

テナント管理者アカウントにアクセスできない場合や、少数のユーザーに同意を限定的にしたい場合は、この方法を使用して単一のユーザーに同意を提供できます。

1. `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code` に移動します。`{application_ID}` は、アプリケーション登録で表示されたアプリケーション ID です。
1. 自分のアカウントでサインインします。
1. アクセス許可を確認してから ［**承諾**］ をクリックします。

ブラウザーは元のアプリへのリダイレクトを試行しますが、そのアプリは実行されていない可能性があります。［**承諾**］ をクリックした後に、「このサイトにアクセスできません」というエラーが表示されることがあります。これは問題ありません。この場合でも同意は記録されています。

## サンプルの実行

> **注**: `WebApplicationInfo` 要素が無効であるという警告またはエラーが Visual Studio で表示されることがあります。このエラーは、ソリューションをビルドするまで表示されない可能性があります。この記事の執筆時点では、Visual Studio では `WebApplicationInfo` を含めるようにスキーマ ファイルが更新されていません。この問題を回避するには、次のリポジトリにある更新されたスキーマ ファイルを使用することができます:[MailAppVersionOverridesV1\_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)
>
> 1. 開発用コンピューターで、既存の MailAppVersionOverridesV1\_1.xsd を見つけます。これは、`./Xml/Schemas/{lcid}` の下の Visual Studio インストール ディレクトリにあるはずです。たとえば、英語 (US) システムでの Visual Stuido 2017 32-ビット版の一般的なインストールでは、完全なパスは `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033` となります。
> 1. 既存のファイルの名前を "`MailAppVersionOverridesV1_1.old`" に変更します。
> 1. このバージョンのファイルをこのリポジトリからフォルダーに移動します。

サンプルを Visual Studio から直接実行できます。**ソリューション エクスプローラー**で \[**AttachmentDemo**] プロジェクトを選択し、(\[プロパティ] ウィンドウの \[**アドイン**] で) 希望する \[**開始動作**] 値を選択します。 インストールされている任意のブラウザーを選択して Outlook on the web を起動することも、**Office デスクトップ クライアント** を選択して Outlook を起動することもできます。**Office デスクトップ クライアント**を選択した場合は、アドインのインストール対象の Office 365 または Outlook.com ユーザーに接続できるように Outlook を構成する必要があります。

> **注:**現時点では、SSO トークン機能は、Windows 版 Outlook 2016 のみでプレビュー中です。

**F5** を押して、プロジェクトをビルドしてデバッグします。ユーザー アカウントとパスワードを求められるはずです。Office 365 テナントまたは Outlook.com アカウントのユーザーを必ず使用してください。アドインがそのユーザー用にインストールされ、Outlook on the web または Outlook が開きます。いずれかのメッセージを選択すると、Outlook のリボンにアドイン ボタンが表示されるはずです。

**デスクトップ版 Outlook のアドイン**

![デスクトップ版 Outlook のリボンにあるアドイン ボタン](readme-images/buttons-outlook.PNG)

**Outlook on the web のアドイン**

![Outlook on the web のアドイン ボタン](readme-images/buttons-owa.PNG)

## 著作権

Copyright (c) Microsoft.All rights reserved.
