# [MOVED] Single Sign-on (SSO) Sample Outlook Add-in

**Note:** This sample was moved to the [PnP-OfficeAddins](https://github.com/OfficeDev/PnP-OfficeAddins) and is located at https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Outlook-Add-in-SSO

This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in.
![image](https://user-images.githubusercontent.com/8559338/121234209-512c7280-c848-11eb-87b2-ae269f39654f.png)

The sample implements an Outlook add-in that uses Office's SSO feature to give the add-in access to Microsoft Graph data. Specifically, it enables the user to save all attachments to their OneDrive. It also shows how to add custom buttons to the Outlook ribbon. The sample illustrates the following concepts:

- [Using the SSO access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) to call the Microsoft Graph API without prompting the user
- If the SSO token is not available, authenticating to the user's OneDrive using the OAuth2 implicit flow via the [office-js-helpers library](https://github.com/OfficeDev/office-js-helpers).
- Using the [Microsoft Graph API](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) to create files in OneDrive.
- Adding [add-in command buttons](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook) to the Outlook ribbon when reading mail, including a UI-less button and a button that opens a task pane
- Implementing a WebAPI to [retrieve attachments via a callback token and the Outlook REST API](https://docs.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item)

## Register the add-in with Azure AD v2.0 endpoint

1. Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.

1. Sign in with the ***admin*** credentials to your Microsoft 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    * Set **Name** to `AttachmentsDemo`.
    * Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    * In the **Redirect URI** section, ensure that **Web** is selected in the drop down and then set the URI to `https://localhost:44349/MessageRead.html`.
        > [!NOTE]
        > The port number in the redirect URI (`44349`) may be different on your development machine. You can find the correct port number for your machine by selecting the **AttachmentDemoWeb** project in **Solution Explorer**, then looking at the **SSL URL** setting in the properties window.
    * Click **Register**.

1. On the **AttachmentsDemo** page, copy and save the values for the **Application (client) ID**. You'll use it in later procedures.

1. Under **Manage**, select **Authentication**. Under **Implicit grant**, turn on the check of **Access tokens**. Then select **Save** button.

1. Under **Manage**, select **Certificates & secrets**. Select the **New client secret** button. Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**. *Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.

1. Under **Manage**, select **Expose an API**. Click the **Set** link.

1. Under **Application ID URI**, change the default value by inserting your host and port number before the GUID listed there. For example, if the default value is `api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`, and your app is running on `localhost:44349`, the value is `api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`. Then click **Save**.

1. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **Scope** name.

1. Set **Who can consent?** to **Admins and users**.

1. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user. Suggestions:

    - **Admin consent title**: Office can act as the user.
    - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.
    - **User consent title**: Office can act as you.
    - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.

1. Ensure that **State** is set to **Enabled**.

1. Select **Add scope** .

1. In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

    For each ID, take these steps:

    a. Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44349/$App ID GUID$/access_as_user`.

    b. Select **Add application**.

1. Under **Manage**, select **API permissions** and then select **Add a permission**. On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search the following permissions.

    * Files.ReadWrite
    * Mail.Read
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > The `User.Read` permission may already be listed by default. It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.

1. Select the check box for each permission as it appears. After selecting the permissions, select the **Add permissions** button at the bottom of the panel.

1. On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Accept** for the confirmation that appears.

    > [!NOTE]
    > After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed. If so, you can start work on the next section, ***but don't forget to come back to the portal and press this button***!

## Configuring the Sample

Before you run the sample, you'll need to do a few things to make it work properly.

Edit [authconfig.js](AttachmentDemoWeb/Scripts/authconfig.js) and replace the `YOUR APP ID HERE` value with the application ID you generated as part of the app registration process.

Edit [AttachmentDemo.xml](AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml) and replace the `YOUR APP ID HERE` value with the application ID you generated as part of the app registration process.

> [!NOTE]
> Make sure that the port number in the `Resource` element matches the port used by your project. It should also match the port you used when registering the application.

Edit [Web.config](AttachmentDemoWeb/Web.config) and replace the `YOUR APP ID HERE` value in with the application ID and `YOUR APP PASSWORD HERE` with the application password you generated as part of the app registration process.

## Provide user consent to the app

If you want to try the add-in using a different tenant than the one where you registered the app, you need to do this step.

You have two choices for providing consent. You can use an administrator account and consent once for all users in your Office 365 tenant, or you can use any account to consent for just that user.

### Provide admin consent for all users

If you have access to a tenant administrator account, this method will allow you to provide consent for all users in your organization, which can be convenient if you have multiple developers that need to develop and test your add-in.

1. Browse to `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your administrator account.
1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

### Provide consent for a single user

If you don't have access to a tenant administrator account, or you just want to limit consent to a few users, this method will allow you to provide consent for a single user.

1. Browse to `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your account.
1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

## Running the Sample

> [!NOTE]
> Visual Studio may show a warning or error about the `WebApplicationInfo` element being invalid. The error may not show up until you try to build the solution. As of this writing, Visual Studio has not updated their schema files to include the `WebApplicationInfo` element. To work around this problem, you can use the updated schema file in this repository: [MailAppVersionOverridesV1_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd).
>
> 1. On your development machine, locate the existing MailAppVersionOverridesV1_1.xsd. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.
> 1. Rename the existing file to `MailAppVersionOverridesV1_1.old`.
> 1. Move the version of the file from this repository into the folder.

> [!NOTE]
> Your browser need to allow third-party cookies because office-js-helpers library uses localStorage. Beware that Google Chrome has a setting which blocks third-party cookies in Incognito mode.

You can run the sample right from Visual Studio. Select the **AttachmentDemo** project in **Solution Explorer**, then choose the **Start Action** value you want (under **Add-in** in the properties window). You can choose any installed browser to launch Outlook on the web, or you can choose **Office Desktop Client** to launch Outlook. If you choose **Office Desktop Client**, be sure to configure Outlook to connect to the Office 365 or Outlook.com user you want to install the add-in for.

Press **F5** to build and debug the project. You should be prompted for a user account and password. Be sure to use a user in your Office 365 tenant, or an Outlook.com account. The add-in will be installed for that user, and either Outlook on the web or Outlook will open. Select any message, and you should see the add-in buttons on the Outlook ribbon.

**Add-in in Outlook on desktop**

![The add-in buttons on the ribbon in Outlook on the desktop](readme-images/buttons-outlook.PNG)

**Add-in in Outlook on the web**

![The add-in buttons in Outlook on the web](readme-images/buttons-owa.PNG)

## Copyright

Copyright (c) Microsoft. All rights reserved.
