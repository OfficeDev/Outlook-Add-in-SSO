using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Xml;

using Newtonsoft.Json;
using Microsoft.Exchange.WebServices.Data;

namespace AttachmentsDemoWeb.Controllers
{
    public class GetAttachmentController : ApiController
    {
        [HttpPost()]
        public string SaveAttachments(AttachmentRequest request)
        {
            try
            {
                // Get the attachments using the EWS Managed API
                Attachment[] attachments = GetAttachmentsWithManagedApi
                    (request.AttachmentIds, request.AuthToken, request.EwsUrl);

                // NOTE: The GetAttachmentsWithSOAP method doesn't use the managed
                // API, and instead just builds the SOAP request manually. This is
                // here to illustrate how this can be done without the managed API.
                //Attachment[] attachments = GetAttachmentsWithSOAP(request.AttachmentIds,
                //    request.AuthToken, request.EwsUrl);

                return SaveAttachments(attachments, request.UserEmail);
            }
            catch (Exception e)
            {
                throw new HttpResponseException(Request.CreateErrorResponse(HttpStatusCode.InternalServerError,
                    "There was an exception: " + e.Message + "\n\n" + e.StackTrace));
            }
        }

        private Attachment[] GetAttachmentsWithManagedApi(string[] attachmentIds, string authToken, string ewsUrl)
        {
            ExchangeService service = new ExchangeService();
            service.Credentials = new OAuthCredentials(authToken);
            service.Url = new Uri(ewsUrl);

            ServiceResponseCollection<GetAttachmentResponse> responses = service.GetAttachments(
                attachmentIds, null, new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.MimeContent));

            List<Attachment> attachments = new List<Attachment>();
            foreach(GetAttachmentResponse response in responses)
            {
                Attachment attachment = new Attachment();
                attachment.AttachmentName = response.Attachment.Name;

                if (response.Attachment is FileAttachment)
                {
                    FileAttachment file = response.Attachment as FileAttachment;
                    attachment.AttachmentBytes = file.Content;
                }
                else if (response.Attachment is ItemAttachment)
                {
                    // Skip item attachments for now
                    // TODO: Add code to extract the MIME content of the item
                    // and build a .EML file to save to OneDrive.
                }

                attachments.Add(attachment);
            }

            return attachments.ToArray();
        }

        private Attachment[] GetAttachmentsWithSOAP(string[] attachmentIds, string authToken, string ewsUrl)
        {
            string getAttachmentRequest =
                @"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
                xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
                xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
                xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
                <soap:Header>
                <t:RequestServerVersion Version=""Exchange2013"" />
                </soap:Header>
                    <soap:Body>
                    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
                    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
                        <AttachmentShape/>
                        <AttachmentIds>";

            foreach (string id in attachmentIds)
            {
                getAttachmentRequest = getAttachmentRequest + @"            <t:AttachmentId Id=""" + id + @"""/>";
            } 
            getAttachmentRequest +=    
                @"        </AttachmentIds>
                    </GetAttachment>
                    </soap:Body>
                </soap:Envelope>";

            // Prepare a web request object.
            HttpWebRequest webRequest = WebRequest.CreateHttp(ewsUrl);
            webRequest.Headers.Add("Authorization", string.Format("Bearer {0}", authToken));
            webRequest.PreAuthenticate = true;
            webRequest.AllowAutoRedirect = false;
            webRequest.Method = "POST";
            webRequest.ContentType = "text/xml; charset=utf-8";

            // Construct the SOAP message for the GetAttchment operation.
            byte[] bodyBytes = System.Text.Encoding.UTF8.GetBytes(getAttachmentRequest);
            webRequest.ContentLength = bodyBytes.Length;

            Stream requestStream = webRequest.GetRequestStream();
            requestStream.Write(bodyBytes, 0, bodyBytes.Length);
            requestStream.Close();

            // Make the request to the Exchange server and get the response.
            HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

            // If the response is okay, create an XML document from the
            // response and process the request.
            if (webResponse.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = webResponse.GetResponseStream();

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(responseStream);

                XmlNodeList fileNameNodes = xmlDocument.GetElementsByTagName("t:Name");
                XmlNodeList byteNodes = xmlDocument.GetElementsByTagName("t:Content");

                Attachment[] attachments = new Attachment[fileNameNodes.Count];

                for (int i = 0; i < fileNameNodes.Count; i++ )
                {
                    attachments[i] = new Attachment();
                    attachments[i].AttachmentName = fileNameNodes[i].InnerText;
                    attachments[i].AttachmentBytes = Convert.FromBase64String(byteNodes[i].InnerText);
                }

                // Close the response stream.
                responseStream.Close();
                webResponse.Close();

                return attachments; //new Attachment() { AttachmentBytes = bytes, AttachmentName = fileName };
            }

            return null;
        }


        private string SaveAttachments(Attachment[] attachments, string userEmail)
        {
            // Get the user's OneDrive endpoint and access token
            OAuthController.OneDriveAccessDetails accessDetails =
                OAuthController.GetUsersOneDriveAccessDetails(userEmail);

            string createAttachmentUri = accessDetails.ApiEndpoint + 
                "/Files/getByPath('{0}')/content?nameConflict=overwrite";

            string returnString = "";

            foreach (Attachment attachment in attachments)
            {
                // Prepare the HTTP request using the new "File" APIs
                HttpWebRequest webRequest =
                    WebRequest.CreateHttp(string.Format(createAttachmentUri, attachment.AttachmentName));
                webRequest.Accept = "application/json";
                webRequest.Headers.Add("Authorization", string.Format("Bearer {0}", accessDetails.AccessToken));
                webRequest.Method = "PUT";
                webRequest.ContentLength = attachment.AttachmentBytes.Length;
                webRequest.ContentType = "application/octet-stream";

                Stream requestStream = webRequest.GetRequestStream();
                requestStream.Write(attachment.AttachmentBytes, 0, attachment.AttachmentBytes.Length);
                requestStream.Close();

                // Make the request to SharePoint and get the response.
                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                if (!string.IsNullOrEmpty(returnString))
                    returnString += "; ";

                // If the response is okay, read it
                if (webResponse.StatusCode == HttpStatusCode.Created)
                {
                    Stream responseStream = webResponse.GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream);

                    returnString += attachment.AttachmentName + ": Success, ID: " + 
                        GetAttachmentIdFromJson(reader.ReadToEnd());
                }
                else
                    returnString += attachment.AttachmentName + ": Error: " + webResponse.StatusCode + Environment.NewLine;
            }

            return returnString;
        }

        private string GetAttachmentIdFromJson(string oneDriveResponse)
        {
            try
            {
                var attachment = JsonConvert.DeserializeObject<dynamic>(oneDriveResponse);
                return attachment.id;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        #region Helper classes
        public class Attachment
        {
            public byte[] AttachmentBytes { get; set; }
            public string AttachmentName { get; set; }
        }

        public class AttachmentRequest
        {
            public string UserEmail { get; set; }
            public string AuthToken { get; set; }
            public string[] AttachmentIds { get; set; }
            public string EwsUrl { get; set; }
        }
        #endregion
    }
}