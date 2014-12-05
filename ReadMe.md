#AttachmentsDemo Sample Mail App
*This is based on the sample mail app originally [shown by Andrew Salamatov at the SharePoint Conference 2014](http://channel9.msdn.com/Events/SharePoint-Conference/2014/SPC391).*

The sample implements a read-mode mail app that activates for items with attachments. It allows the user to save all attachments to their OneDrive for Business. The sample illustrates the following concepts:
 
- Implementing a [read-mode mail app](http://msdn.microsoft.com/en-us/library/office/fp161135(v=office.15).aspx)
- Implementing a WebAPI to [retrieve attachments via a callback token](http://msdn.microsoft.com/en-us/library/office/dn148008(v=office.15).aspx)
- Using the [Discovery Service](http://msdn.microsoft.com/en-us/office/office365/api/discovery-service-rest-operations) to find a user's OneDrive endpoint
- Using the [Files REST API](http://msdn.microsoft.com/en-us/office/office365/api/files-rest-operations) to create files in OneDrive for Business

##Configuring the Sample##

Before you run the sample, you'll need to do a few things to make it work properly.

1. You need an Office 365 tenant. While mail apps will work with on-premise installations of Exchange, the Files REST API requires Office 365. If you don't already have an Office 365 tenant, you can get an Office 365 Developer Subscription, either through an existing MSDN subscription, or via a [free trial](https://portal.microsoftonline.com/Signup/MainSignUp.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK). Lots of good information on this [here](http://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment).
2. You need to register the sample application in your tenant's Azure Active Directory in order to obtain a client ID and client secret for accessing the Files REST API. There's a walkthrough of adding an application via the Azure Management Portal [here](http://msdn.microsoft.com/en-us/library/azure/dn132599.aspx). The important values are:

- Name: AttachmentsDemo
- Type: Web Application and/or Web API
- Sign-on URL: The SSL URL of the AttachmentsDemoWeb project in the solution. For example, if you're running it in Visual Studio, it is probably something like "https://localhost:44307".
- App ID URI: Same as Sign-on URL.
- Permissions to other Applications: Office 365 SharePoint Online, set Delegated Permissions to enable "Edit or delete users' files".

3. Get the client ID from the app's registration in the Azure Management Portal, and generate a key for the app. This is done on the "Configure" tab of the app in the portal. Copy the client ID into the **ClientId** variable, and copy the key into the **ClientSecret** variable. These are both found in OAuthController.cs.

##Running the Sample##

You can run the sample right from Visual Studio. You should be prompted for a user account and password. Be sure to use a user in your Office 365 tenant. The mail app will be installed for that user, and Outlook Web Access will open. Select any message with file attachments, and you should see an **AttachmentsDemo** button in the reading pane.