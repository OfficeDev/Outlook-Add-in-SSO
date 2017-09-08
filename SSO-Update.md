# Updating the add-in to use SSO

## Update the app registration

Because the current version uses Graph to write files to the user's OneDrive, we already have an app registration. We need to update it to meet the requirements of SSO.

### Add Web API details

1. Go to https://apps.dev.microsoft.com and edit the existing app registration.
1. Click **Add Platform**. Choose **Web API**.
1. Under **Application ID URI**, change the default value by inserting your host and port number before the GUID listed there. For example:
    1. Let's say the default value is `api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`.
    1. Look at the value of **Redirect URLs** under the existing **Web** platform entry. Let's say it is `https://localhost:44349/MessageRead.html`. Copy the `localhost:44349` part.
    1. Insert that into the default **Application ID URI** value after the `//`, then add another `/` after it, so the result is `api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`
1. Under **Pre-authorized applications**, enter `d3590ed6-52b3-4102-aeff-aad2292ab01c` for the **Application ID**. Click the **Scope** dropdown and select the only entry there. This preauthorizes Office to access the app.

When you're done, the **Web API** section should look similar to the following:

![Screenshot of Web API platform section in app registration](readme-images/web-api-app-registration.PNG)

### Add Microsoft Graph Permissions

1. Locate the **Microsoft Graph Permissions** section in the app registration. Next to **Delegated Permissions**, click **Add**.
1. Select **Files.ReadWrite**, **Mail.Read**, and **profile**. Click **OK**.

When you're done, the **Microsoft Graph Permissions** section should looke like the following:

![Screenshot of Microsoft Graph Permissions section in app registration](readme-images/graph-permissions-app-registration.PNG)

### Generate an application secret

1. Locate the **Application Secrets** section in the app registration. Click **Generate New Password**.
1. Copy the password that is generated and save it somewhere safe. We'll need this in a bit. Click **OK**.

### Commit changes

Scroll down to the bottom of the app registration and click **Save**.

## Provide user consent to the app

In this step we will provide user consent to the permissions we just configured on the app. This step is **only** necessary because we will be sideloading the add-in for development and testing. Normally a production add-in will be listed in the Office Store, and users will be prompted to give consent during the installation process through the store.

You have two choices for providing consent. You can use an administrator account and consent once for all users in your Office 365 organization, or you can use any account to consent for just that user.

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

## Update the add-in manifest

The next step is to update the add-in manifest to enable the SSO feature.

1. Open the `./AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml` file.
1. Copy the entire `VersionOverrides` element. We are going to use this as the basis for a second `VersionOverrides` element that uses the `VersionOverridesV1_1` schema.
1. Paste the copied data into the manifest just before the closing tag for the existing `VersionOverrides` element.

The following steps all apply to the newly inserted `VersionOverrides` element.

1. In the `VersionOverrides` element, change the following attributes:
    - Change `xmlns` to `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`
    - Change `xsi:type` to `VersionOverridesV1_1`
1. After the `Resources` element, insert the following XML, replacing `YOUR APP ID HERE` with the application ID from your app registration:

    ```xml
    <WebApplicationInfo>
      <Id>YOUR APP ID HERE</Id>
      <Resource>api://localhost:44349/YOUR APP ID HERE</Resource>
      <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.readwrite</Scope>
        <Scope>mail.read</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

    > **Note**: Make sure that the port number in the `Resource` element matches the port used by your project. It should also match the port you used when registering the application.
1. Save your changes.

Also see this [commit](https://github.com/OfficeDev/outlook-add-in-attachments-demo/commit/318e55e5b613ef1aec9b1a8fbc9335bd1cab6a65) on GitHub for the specific change to the manifest.

> **Note**: At this point, Visual Studio may show a warning or error about the `WebApplicationInfo` element being invalid. The error may not show up until you try to build the solution. As of this writing, Visual Studio has not updated their schema files to include the `WebApplicationInfo` element. To work around this problem, you can use the updated schema file in this repository: [MailAppVersionOverridesV1_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd).
>
> 1. On your development machine, locate the existing MailAppVersionOverridesV1_1.xsd. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.
> 1. Rename the existing file to `MailAppVersionOverridesV1_1.old`.
> 1. Move the version of the file from this repository into the folder.

## Update the Web API to handle On-behalf-of flow

Next we need to update the Web API. Currently it is very simple, requiring the add-in to send it both an Outlook token (which it gets from `getCallbackTokenAsync` and a Graph token for OneDrive access (which it gets from OAuth)). We need to update it to accept a bearer token in the `Authorization` header and use that token in the on-behalf-of flow to get an access token.

### Install NuGet packages

1. In Visual Studio, click **Tools**, **NuGet Package Manager**, **Manage NuGet Packages for Solution...**.
1. Click **Browse**, then search for `System.IdentityModel.Tokens.Jwt`.
1. Select select the result, then make sure that the **AttachmentDemoWeb** project is selected. In the **Version** dropdown, select the latest `4.0.*` build and click **Install**.

Repeat those steps (without worrying about specific version) to install these additional packages:

- `Microsoft.Owin`
- `Microsoft.Owin.Security`
- `Microsoft.Owin.Security.Jwt`
- `Microsoft.Owin.Security.OpenIdConnect`
- `Microsoft.Owin.Host.SystemWeb`

Finally, install the Microsoft Authentication Library (MSAL). This is a preview package, so to find it you must select the **Include prerelease** checkbox next to the search box, then search for `Microsoft Identity Client`.

### Configure OWIN middleware

We're going to use OWIN to handle parsing and validating the access token. In order to get that to work, we need to setup the proper OWIN middleware using an OWIN startup class. We'll also add a custom token provider class that can get the signing tokens from the Azure endpoints to perform validation.

#### Add the custom token provider

1. In Solution Explorer, right-click the **App_Start** folder and choose **Add**, then **Class...**. Name the class `OpenIdConnectCachingSecurityTokenProvider` and click **Add**.
1. Open the `OpenIdConnectCachingSecurityTokenProvider.cs` file and replace its entire contents with the following code.

    ```csharp
    using Microsoft.IdentityModel.Protocols;
    using Microsoft.Owin.Security.Jwt;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens;
    using System.Threading;

    namespace AttachmentDemoWeb.App_Start
    {
        public class OpenIdConnectCachingSecurityTokenProvider : IIssuerSecurityTokenProvider
        {
            public ConfigurationManager<OpenIdConnectConfiguration> _configManager;
            private string _issuer;
            private IEnumerable<SecurityToken> _tokens;
            private readonly string _metadataEndpoint;

            private readonly ReaderWriterLockSlim _synclock = new ReaderWriterLockSlim();

            public OpenIdConnectCachingSecurityTokenProvider(string metadataEndpoint)
            {
                _metadataEndpoint = metadataEndpoint;
                _configManager = new ConfigurationManager<OpenIdConnectConfiguration>(metadataEndpoint);

                RetrieveMetadata();
            }

            /// <summary>
            /// Gets the issuer the credentials are for.
            /// </summary>
            /// <value>
            /// The issuer the credentials are for.
            /// </value>
            public string Issuer
            {
                get
                {
                    RetrieveMetadata();
                    _synclock.EnterReadLock();
                    try
                    {
                        return _issuer;
                    }
                    finally
                    {
                        _synclock.ExitReadLock();
                    }
                }
            }

            /// <summary>
            /// Gets all known security tokens.
            /// </summary>
            /// <value>
            /// All known security tokens.
            /// </value>
            public IEnumerable<SecurityToken> SecurityTokens
            {
                get
                {
                    RetrieveMetadata();
                    _synclock.EnterReadLock();
                    try
                    {
                        return _tokens;
                    }
                    finally
                    {
                        _synclock.ExitReadLock();
                    }
                }
            }

            private void RetrieveMetadata()
            {
                _synclock.EnterWriteLock();
                try
                {
                    OpenIdConnectConfiguration config = _configManager.GetConfigurationAsync().Result;
                    _issuer = config.Issuer;
                    _tokens = config.SigningTokens;
                }
                finally
                {
                    _synclock.ExitWriteLock();
                }
            }
        }
    }
    ```

#### Add the OWIN startup class

1. In Solution Explorer, right-click the **AttachmentDemoWeb** project and choose **Add**, **New Item...**. Search for `OWIN` and select **OWIN Startup class**. Name the file `Startup.cs` and click **Add**.
1. Open the `Startup.cs` file and replace its entire contents with the following code.

    ```csharp
    using AttachmentDemoWeb.App_Start;
    using Microsoft.Owin;
    using Microsoft.Owin.Security.Jwt;
    using Microsoft.Owin.Security.OAuth;
    using Owin;
    using System.IdentityModel.Tokens;

    [assembly: OwinStartup(typeof(AttachmentDemoWeb.Startup))]

    namespace AttachmentDemoWeb
    {
        public class Startup
        {
            public void Configuration(IAppBuilder app)
            {
                // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=316888
                var tokenValidationParms = new TokenValidationParameters
                {
                    // Audience MUST be the application ID of the app
                    ValidAudience = "YOUR APP ID HERE",
                    // Since this is multi-tenant we will validate the issuer in the controller
                    ValidateIssuer = false,
                    SaveSigninToken = true
                };

                app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
                {
                    AccessTokenFormat = new JwtFormat(tokenValidationParms,
                        new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
                });
            }
        }
    }
    ```
1. Replace `YOUR APP ID HERE` with the application ID from your app registration.