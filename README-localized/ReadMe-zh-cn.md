---
languages:
- javascript
page_type: sample
description: "本示例实现可将按钮添加到 Outlook 功能区的 Outlook 加载项。"
products:
- office
- office-outlook
urlFragment: outlook-ribbon-addin
---

# AttachmentsDemo 示例 Outlook 加载项

本示例实现可将按钮添加到 Outlook 功能区的 Outlook 加载项。用户可以通过它将所有附件保存到他们的 OneDrive。示例将说明下列概念：
 
- 阅读邮件时向 Outlook 功能区添加“[加载项命令按钮](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)”，包括无用户界面的按钮和打开任务窗格的按钮
- 实现 WebAPI，以“[通过回调令牌和 Outlook REST API 检索附件](https://dev.office.com/docs/add-ins/outlook/use-rest-api) ”
- [使用 SSO 访问令牌](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)在不提示用户的情况下调用 Microsoft Graph API
- 如果 SSO 令牌不可用，通过“[office-js-helpers 库](https://github.com/OfficeDev/office-js-helpers) ”对使用 OAuth2 隐式流的用户 OneDrive 进行身份验证。
- 使用 [Microsoft Graph API](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) 在 OneDrive 中创建文件。

## 配置示例

运行示例前，需要执行一些操作才能使其正常工作。

1. 需要 Office 365 租户或 Outlook.com 帐户。虽然邮件应用程序使用 Exchange 的本地安装，Microsoft Graph API 仍需要 Office 365 或 Outlook.com。
2. 获得访问 Microsoft Graph API 的应用程序 ID 需要，在 [Microsoft 应用程序注册门户](https://apps.dev.microsoft.com)中注册示例应用程序。
    1. 浏览到 [Microsoft 应用程序注册门户](https://apps.dev.microsoft.com)。如果未要求登录，单击“**转至应用程序列表**”按钮并使用 Microsoft 帐户 (Outlook.com) /工作或学校帐户 (Office 365) 进行登录。登录后，单击“**添加应用程序**”按钮。在名称中输入`AttachmentsDemo`并点击**创建应用程序**。
    1. 找到“**应用程序机密**”部分，并单击“**生成新密码**”按钮。含有生成密码的对话框显示。关闭对话框并保存前，复制此数值。
    1. 找到“**平台**”部分，再单击“**添加平台**”。选择**Web**，随后在“**重定向 URI** 下输入 `https://localhost:44349/MessageRead.html`。
        > **注意：**重定向 URI 中的端口号（`44349`）可能在开发计算机上不同。可通过在“**解决方案资源管理器**”中选择 “**AttachmentDemoWeb**” 项目查找计算机正确的端口后，然后在“属性” 窗口中的“**开发服务器**”下查看 **SSL URL** 设置。
        
    1. 单击“**添加平台**”。选择“**Web API**”。按照下列方式配置此部分：
        - 在“**应用程序 ID URI**”下，通过在列出的 GUID 前插入主机和端口号，更改默认值。例如，如果默认值是 `api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`，应用程序在 `localhost:44349` 上运行，则数值是 `api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`。
        - 在“**预授权应用程序**”，为**应用程序 ID**输入`d3590ed6-52b3-4102-aeff-aad2292ab01c`。单击“**范围**”下拉列表，并只选择项。此操作将预授权桌面版 Office（Windows）应用程序访问应用程序。
        - 在 "**预授权的应用程序**”下，输入 **应用程序 ID**的 `bc59ab01-8403-45c6-8796-ac3ef710b3e3`。单击“**范围**”下拉列表，并只选择项。此操作会预授权 Outlook 网页版访问应用程序。
    1. 在应用程序注册中找到“**Microsoft Graph 权限**”部分。在“**委派权限**”旁，单击“**添加**”。选择“**Files.ReadWrite**”、“**Mail.Read**”、“**offline\_access**”、“**openid**”和“**profile**”。单击**“确定”**。

单击“**保存**”，完成注册。复制**应用程序 ID**并复制到之前保存到应用程序密码的相同位置。我们很快就会需要这些值。

完成上述步骤后，应用程序的注册详细信息应如下所示。

![完成的应用程序注册](readme-images/app-registration.PNG)
![完成的应用程序第 2 部分](readme-images/web-api-app-registration.PNG)

编辑 [authconfig.js](AttachmentDemoWeb/Scripts/authconfig.js) 并将 `YOUR APP ID HERE` 值替换成在应用程序注册过程中生成的应用 ID。

编辑 [AttachmentDemo.xml](AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml) 并将 `YOUR APP ID HERE` 值替换成在应用程序注册过程中生成的应用程序 ID。

> **注意**：确保`资源`元素中的端口号与项目中使用的端口号匹配。注册应用程序时，还应与所使用的端口匹配。

编辑 [authconfig.js](AttachmentDemoWeb/Web.config)，将 `YOUR APP ID HERE` 值替换成应用程序 ID ，将 `YOUR APP PASSWORD HERE` 值替换为作为应用程序注册过程中生成的应用程序密码。

## 为应用程序提供用户授权

在此步骤中，我们将提供在应用程序上配置的权限用户授权。只有在针对开发和测试旁加载加载项时，**才**需要此步骤。通常情况下，Office 应用商店中将列出一个生产加载项，并且系统将提示用户在安装过程中通过应用商店提供授权。

有两种提供授权的选项。可以使用管理员帐户一次性地向 Office 365 组织中的所有用户授权，也可以使用任何帐户仅向相应的用户授权。

### 为所有用户提供管理员授权

如果具有租户管理员帐户的访问权限，则可以使用此方法为组织中的所有用户提供授权，如果有多个开发人员需要开发和测试你的加载项，这一方法则很便捷。

1. 浏览到 `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`，其中 `{application_ID}` 是显示在应用程序注册中的应用程序 ID。
1. 使用管理员帐户 1 登录。审阅权限并单击“**接受**”。

浏览器将尝试重定向回你的应用，该应用可能没有运行。单击“**接受**”后，可能会看到“该网站无法访问”的错误。这不会有影响，授权仍会被记录。

### 为单个用户提供授权

如果没有租户管理员帐户的访问权限，或只想将授权限制在几个用户内，可以使用此方法为单个用户提供授权。

1. 浏览到 `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code`，其中 `{application_ID}` 是显示在应用程序注册中的应用程序 ID。
1. 使用你的帐户登录。
1. 审阅权限并单击“**接受**”。

浏览器将尝试重定向回你的应用，该应用可能没有运行。单击“**接受**”后，可能会看到“该网站无法访问”的错误。这不会有影响，授权仍会被记录。

## 运行示例

> **注意**：Visual Studio 可能显示有关 `WebApplicationInfo` 元素无效的警告或错误。生成解决方案前，该错误不会显示。截至到该写入操作，Visual Studio 尚未更新架构文件以包含 `WebApplicationInfo` 元素。若要解决此问题，可使用该存储库中更新的架构文件：[MailAppVersionOverridesV1\_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)。
>
> 1. 在开发计算机上找到现有的 MailAppVersionOverridesV1\_1.xsd。它应位于 `./Xml/Schemas/{lcid}`下的 Visual Studio 安装目录中。例如在 VS 2017 32 位的典型安装（英语（美国））系统中，完整的路径应为 `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`。
> 1. 重命名现有文件为 `MailAppVersionOverridesV1_1.old`。
> 1. 将此存储库中的文件版本移动至文件夹中。

可直接从 Visual Studio 运行示例。在“**解决方案资源管理器**”中选择 “**AttachmentDemo**” 项目，随后选择期望的“**开始操作**”值（在属性窗口中的“**加载项**”下）。可选择任何已安装的浏览器启动 Outlook 网页版，或者选择“**Office 桌面客户端**”启动 Outlook。如果选择 “**Office 桌面版客户端**”，确保配置 Outlook 以连接至 想要安装加载项的 Office 365 或 Outlook.com 用户。

> **注意：**SSO 令牌功能目前仅供 Outlook 2016 for Windows 预览版可用。

按下 **F5** 生成并调试项目。系统提示输入用户账户和密码。请务必使用 Office 365 租户中的用户或 Outlook.com 帐户。加载项将为该用户安装，并打开 Outlook 网页版或 Outlook。选择任一邮件，将在 Outlook 功能区上看到“加载项”按钮。

**Outlook 桌面版加载项**

![Outlook 网页版功能区上的加载项按钮](readme-images/buttons-outlook.PNG)

**Outlook 网页版加载项**

![Outlook 网页版加载项按钮](readme-images/buttons-owa.PNG)

## 版权信息

版权所有 (c) Microsoft。保留所有权利。
