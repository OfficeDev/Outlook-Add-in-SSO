// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using AttachmentDemoWeb.Models;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;

namespace AttachmentDemoWeb.Controllers
{
    public class SaveAttachmentsController : ApiController
    {
        // POST api/<controller>
        public async Task<IHttpActionResult> Post([FromBody]SaveAttachmentRequest request)
        {
            if (Request.Headers.Contains("Authorization"))
            {
                // Request contains bearer token, validate it
                var scopeClaim = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope");
                if (scopeClaim != null)
                {
                    // Check the allowed scopes
                    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
                    if (!addinScopes.Contains("access_as_user"))
                    {
                        return BadRequest("The bearer token is missing the required scope.");
                    }
                }
                else
                {
                    return BadRequest("The bearer token is invalid.");
                }

                // Passed validation, process the request
                return await SaveAttachmentsWithSsoToken(request);
            }
            else
            {
                // No bearer token, so this is a request without SSO
                // Access tokens are included in the request
                return await SaveAttachmentsWithDistinctTokens(request);
            }
        }

        private async Task<IHttpActionResult> SaveAttachmentsWithSsoToken(SaveAttachmentRequest request)
        {
            // First retrieve the raw access token
            var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
            if (bootstrapContext != null)
            {
                // Use MSAL to invoke the on-behalf-of flow to exchange token for a Graph token
                UserAssertion userAssertion = new UserAssertion(bootstrapContext);
                string[] graphScopes = { "Files.ReadWrite", "Mail.Read" };
                var client = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:AppId"])
                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                    .WithClientSecret(ConfigurationManager.AppSettings["ida:AppPassword"])
                    .WithRedirectUri(ConfigurationManager.AppSettings["ida:RedirectUri"])
                    .Build();

                AuthenticationResult authResult = await client.AcquireTokenOnBehalfOf(graphScopes, userAssertion).ExecuteAsync();

                // Initialize a Graph client
                GraphServiceClient graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) => {
                            // Add the OneDrive access token to each outgoing request
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                            return Task.FromResult(0);
                        }));

                foreach (string attachmentId in request.attachmentIds)
                {
                    var attachment = await graphClient.Me.Messages[request.messageId].Attachments[attachmentId].Request().GetAsync();

                    // Is this a file or an Outlook item?
                    if (string.Compare(attachment.ODataType, "#microsoft.graph.itemAttachment") == 0)
                    {
                        // Re-request the attachment with the item expanded
                        var itemAttachment = await graphClient.Me.Messages[request.messageId].Attachments[attachmentId].Request()
                            .Expand("microsoft.graph.itemAttachment/item").GetAsync() as ItemAttachment;

                        // Serialize the item to JSON and save to OneDrive
                        string jsonItem = JsonConvert.SerializeObject(itemAttachment.Item);
                        MemoryStream fileStream = new MemoryStream();
                        StreamWriter sw = new StreamWriter(fileStream);
                        sw.Write(jsonItem);
                        sw.Flush();
                        fileStream.Position = 0;
                        bool success = await SaveFileToOneDrive(graphClient, itemAttachment.Name + ".json", fileStream);
                        if (!success)
                        {
                            return BadRequest(string.Format("Could not save {0} to OneDrive", itemAttachment.Name));
                        }
                    }
                    else
                    {
                        var fileAttachment = attachment as FileAttachment;

                        // For files, we can build a stream directly from ContentBytes
                        if (fileAttachment.Size < (4 * 1024 * 1024))
                        {
                            MemoryStream fileStream = new MemoryStream(fileAttachment.ContentBytes);
                            bool success = await SaveFileToOneDrive(graphClient, fileAttachment.Name, fileStream);
                            if (!success)
                            {
                                return BadRequest(string.Format("Could not save {0} to OneDrive", fileAttachment.Name));
                            }
                        }
                        else
                        {
                            // TODO: Add code here to handle larger files. See:
                            // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_createuploadsession
                            // and
                            // https://github.com/microsoftgraph/aspnet-snippets-sample/blob/master/Graph-ASPNET-46-Snippets/Microsoft%20Graph%20ASPNET%20Snippets/Models/FilesService.cs
                            return BadRequest("File is too large for simple upload.");
                        }
                    }
                }

                return Ok();
            }
            else
            {
                return BadRequest("Could not retrieve access token from request.");
            }
        }

        private async Task<IHttpActionResult> SaveAttachmentsWithDistinctTokens(SaveAttachmentRequest request)
        {
            // Validate request
            if (request == null || !request.IsValid())
            {
                return BadRequest("One or more parameters is missing.");
            }

            // Initialize a Graph client
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) => {
                        // Add the OneDrive access token to each outgoing request
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", request.oneDriveToken);
                        return Task.FromResult(0);
                    }));

            // First get the attachments from the message
            // The token we get from the add-in is valid for the Outlook API,
            // not the Microsoft Graph API. That means we need to invoke the
            // Outlook endpoint directly, and we can't use the Graph library.
            // Build the request URI to the message attachments collection
            string baseAttachmentsUri = request.outlookRestUrl;
            if (!baseAttachmentsUri.EndsWith("/"))
                baseAttachmentsUri += "/";
            baseAttachmentsUri += "v2.0/me/messages/" + request.messageId + "/attachments/";

            using (var client = new HttpClient())
            {
                foreach (string attachmentId in request.attachmentIds)
                {
                    var getAttachmentReq = new HttpRequestMessage(HttpMethod.Get, baseAttachmentsUri + attachmentId);

                    // Headers
                    getAttachmentReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", request.outlookToken);
                    getAttachmentReq.Headers.UserAgent.Add(new ProductInfoHeaderValue("AttachmentsDemoOutlookAddin", "1.0"));
                    getAttachmentReq.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var result = await client.SendAsync(getAttachmentReq);

                    string json = await result.Content.ReadAsStringAsync();
                    OutlookAttachment attachment = JsonConvert.DeserializeObject<OutlookAttachment>(json);

                    // Is this a file or an Outlook item?
                    if (attachment.Type.ToLower().Contains("itemattachment"))
                    {
                        // Currently REST API doesn't support access to the MIME stream
                        // So for now, just get the JSON representation of the attached item and save it
                        // as a JSON file
                        var getAttachedItemJsonReq = new HttpRequestMessage(HttpMethod.Get, baseAttachmentsUri +
                            attachmentId + "?$expand=Microsoft.OutlookServices.ItemAttachment/Item");

                        getAttachedItemJsonReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", request.outlookToken);
                        getAttachedItemJsonReq.Headers.UserAgent.Add(new ProductInfoHeaderValue("AttachmentsDemoOutlookAddin", "1.0"));
                        getAttachedItemJsonReq.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                        var getAttachedItemResult = await client.SendAsync(getAttachedItemJsonReq);

                        Stream jsonAttachedItem = await getAttachedItemResult.Content.ReadAsStreamAsync();
                        bool success = await SaveFileToOneDrive(graphClient, attachment.Name + ".json", jsonAttachedItem);
                        if (!success)
                        {
                            return BadRequest(string.Format("Could not save {0} to OneDrive", attachment.Name));
                        }
                    }
                    else
                    {
                        // For files, we can build a stream directly from ContentBytes
                        if (attachment.Size < (4 * 1024 * 1024))
                        {
                            MemoryStream fileStream = new MemoryStream(Convert.FromBase64String(attachment.ContentBytes));
                            bool success = await SaveFileToOneDrive(graphClient, attachment.Name, fileStream);
                            if (!success)
                            {
                                return BadRequest(string.Format("Could not save {0} to OneDrive", attachment.Name));
                            }
                        }
                        else
                        {
                            // TODO: Add code here to handle larger files. See:
                            // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_createuploadsession
                            // and
                            // https://github.com/microsoftgraph/aspnet-snippets-sample/blob/master/Graph-ASPNET-46-Snippets/Microsoft%20Graph%20ASPNET%20Snippets/Models/FilesService.cs
                            return BadRequest("File is too large for simple upload.");
                        }
                    }
                }
            }
            return Ok();
        }

        private async Task<bool> SaveFileToOneDrive(GraphServiceClient client, string fileName, Stream fileContent)
        {
            string relativeFilePath = "Outlook Attachments/" + MakeFileNameValid(fileName);

            try
            {
                // This method only supports files 4MB or less
                DriveItem newItem = await client.Me.Drive.Root.ItemWithPath(relativeFilePath)
                    .Content.Request().PutAsync<DriveItem>(fileContent);
            }
            catch (ServiceException)
            {
                return false;
            }

            return true;
        }

        private string MakeFileNameValid(string originalFileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", originalFileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
        }
    }
}