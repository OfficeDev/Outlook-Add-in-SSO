---
languages:
- javascript
page_type: sample
description: "The sample implements an Outlook add-in that adds buttons to the Outlook ribbon."
products:
- office
- office-outlook
urlFragment: outlook-ribbon-addin
---

# AttachmentsDemo Sample Outlook Add-in

The sample implements an Outlook add-in that adds buttons to the Outlook ribbon. It allows the user to save all attachments to their OneDrive. The sample illustrates the following concepts:
 
- Adding [add-in command buttons](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook) to the Outlook ribbon when reading mail, including a UI-less button and a button that opens a task pane
- Implementing a WebAPI to [retrieve attachments via a callback token and the Outlook REST API](https://dev.office.com/docs/add-ins/outlook/use-rest-api)
- [Using the SSO access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) to call the Microsoft Graph API without prompting the user
- If the SSO token is not available, authenticating to the user's OneDrive using the OAuth2 implicit flow via the [office-js-helpers library](https://github.com/OfficeDev/office-js-helpers).
- Using the [Microsoft Graph API](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) to create files in OneDrive.

## Configuring the Sample

Before you run the sample, you'll need to do a few things to make it work properly.

1. You need an Office 365 tenant or Outlook.com account. While mail apps will work with on-premise installations of Exchange, the Microsoft Graph API requires Office 365 or Outlook.com.
2. You need to register the sample application in the [Microsoft Application Registration Portal](https://apps.dev.microsoft.com) in order to obtain an app ID for accessing the Microsoft Graph API.
    1. Browse to the [Microsoft Application Registration Portal](https://apps.dev.microsoft.com). If you're not asked to sign in, click the **Go to app list** button and sign in with either your Microsoft account (Outlook.com), or your work or school account (Office 365). Once you're signed in, click the **Add an app** button. Enter `AttachmentsDemo` for the name and click **Create application**.
    1. Locate the **Application Secrets** section, and click **Generate New Password**. A dialog box will appear with the generated password. Copy this value before dismissing the dialog and save it.
    1. Locate the **Platforms** section, and click **Add Platform**. Choose **Web**, then enter `https://localhost:44349/MessageRead.html` under **Redirect URIs**.
        > **Note:** The port number in the redirect URI (`44349`) may be different on your development machine. You can find the correct port number for your machine by selecting the **AttachmentDemoWeb** project in **Solution Explorer**, then looking at the **SSL URL** setting under **Development Server** in the properties window.
        
    1. Click **Add Platform**. Choose **Web API**. Configure this section as follows:
        - Under **Application ID URI**, change the default value by inserting your host and port number before the GUID listed there. For example, if the default value is `api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`, and your app is running on `localhost:44349`, the value is `api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`.
        - Under **Pre-authorized applications**, enter `d3590ed6-52b3-4102-aeff-aad2292ab01c` for the **Application ID**. Click the **Scope** dropdown and select the only entry there. This pre-authorizes Desktop Office (on Windows) to access the app.
        - Under **Pre-authorized applications**, enter `bc59ab01-8403-45c6-8796-ac3ef710b3e3` for the **Application ID**. Click the **Scope** dropdown and select the only entry there. This pre-authorizes Outlook on the web to access the app.
    1. Locate the **Microsoft Graph Permissions** section in the app registration. Next to **Delegated Permissions**, click **Add**. Select **Files.ReadWrite**, **Mail.Read**, **offline_access**, **openid**, and **profile**. Click **OK**.

Click **Save** to complete the registration. Copy the **Application Id** and save it in the same place with the app password you saved earlier. We'll need those values soon.

Here's what the details of your app registration should look like when you are done.

![The completed app registration](readme-images/app-registration.PNG)
![The completed app registration part 2](readme-images/web-api-app-registration.PNG)

Edit [authconfig.js](AttachmentDemoWeb/Scripts/authconfig.js) and replace the `YOUR APP ID HERE` value with the application ID you generated as part of the app registration process.

Edit [AttachmentDemo.xml](AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml) and replace the `YOUR APP ID HERE` value with the application ID you generated as part of the app registration process.

> **Note**: Make sure that the port number in the `Resource` element matches the port used by your project. It should also match the port you used when registering the application.

Edit [Web.config](AttachmentDemoWeb/Web.config) and replace the `YOUR APP ID HERE` value in  with the application ID and `YOUR APP PASSWORD HERE` with the application password you generated as part of the app registration process.

## Provide user consent to the app

In this step we will provide user consent to the permissions we just configured on the app. This step is **only** necessary because we will be side-loading the add-in for development and testing. Normally a production add-in will be listed in the Office Store, and users will be prompted to give consent during the installation process through the store.

You have two choices for providing consent. You can use an administrator account and consent once for all users in your Office 365 organization, or you can use any account to consent for just that user.

### Provide admin consent for all users

If you have access to a tenant administrator account, this method will allow you to provide consent for all users in your organization, which can be convenient if you have multiple developers that need to develop and test your add-in.

1. Browse to `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your administrator account.1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

### Provide consent for a single user

If you don't have access to a tenant administrator account, or you just want to limit consent to a few users, this method will allow you to provide consent for a single user.

1. Browse to `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your account.
1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

## Running the Sample

> **Note**: Visual Studio may show a warning or error about the `WebApplicationInfo` element being invalid. The error may not show up until you try to build the solution. As of this writing, Visual Studio has not updated their schema files to include the `WebApplicationInfo` element. To work around this problem, you can use the updated schema file in this repository: [MailAppVersionOverridesV1_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd).
>
> 1. On your development machine, locate the existing MailAppVersionOverridesV1_1.xsd. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.
> 1. Rename the existing file to `MailAppVersionOverridesV1_1.old`.
> 1. Move the version of the file from this repository into the folder.

You can run the sample right from Visual Studio. Select the **AttachmentDemo** project in **Solution Explorer**, then choose the **Start Action** value you want (under **Add-in** in the properties window). You can choose any installed browser to launch Outlook on the web, or you can choose **Office Desktop Client** to launch Outlook. If you choose **Office Desktop Client**, be sure to configure Outlook to connect to the Office 365 or Outlook.com user you want to install the add-in for.

> **Note:** The SSO token feature is in preview in Outlook 2016 for Windows only at this time.

Press **F5** to build and debug the project. You should be prompted for a user account and password. Be sure to use a user in your Office 365 tenant, or an Outlook.com account. The add-in will be installed for that user, and either Outlook on the web or Outlook will open. Select any message, and you should see the add-in buttons on the Outlook ribbon.

**Add-in in Outlook on desktop**

![The add-in buttons on the ribbon in Outlook on the desktop](readme-images/buttons-outlook.PNG)

**Add-in in Outlook on the web**

![The add-in buttons in Outlook on the web](readme-images/buttons-owa.PNG)

## Copyright

Copyright (c) Microsoft. All rights reserved.
