// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web.Http;

using Newtonsoft.Json;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace AttachmentsDemoWeb.Controllers
{
    public class OAuthController : ApiController
    {
        // Register the app in Azure AD to get these values.
        // Client ID is found on the "Configure" tab for the application in Azure Management Portal
        private static readonly string ClientId = "";
        // Client Secret is generated on the "Configure" tab for the application in Azure Management Portal,
        // Under "Keys"
        private static readonly string ClientSecret = "";
        
        // OAuth endpoints
        private const string OAuthUrl = "https://login.windows.net/{0}";
        private static readonly string AuthorizeUrlNoResource = string.Format(CultureInfo.InvariantCulture,
            OAuthUrl,
            "common/oauth2/authorize?response_type=code&client_id={0}&redirect_uri={1}&state={2}");
        private static readonly Uri RedirectUrl = new Uri(
            System.Web.HttpContext.Current.Request.Url, "/AppRead/OAuthRedirect.html");

        // Discovery service constants
        private const string DiscoveryResource = "https://api.office.com/discovery/";
        private const string DiscoveryUrl = "https://api.office.com/discovery/v1.0/me/services";
        private const string OneDriveCapability = "MyFiles";

        [HttpPost()]
        public bool IsConsentInPlace(AuthorizationRequest request)
        {
            Storage.AppConfig config = Storage.AppConfigCache.GetUserConfig(request.UserEmail);
            
            // If we have a refresh token for this user, we already have consent
            if (config != null && !string.IsNullOrEmpty(config.RefreshToken))
            {
                return true;
            }
            return false;
        }

        [HttpPost()]
        public string GetAuthorizationUrl(AuthorizationRequest request)
        {
            // Generate a new GUID to add to the request.
            // Save the GUID mapped to the user, so we can look up the user
            // once we have the auth response.
            string stateGuid = Guid.NewGuid().ToString();
            Storage.AppConfigCache.AddStateGuid(stateGuid, request.UserEmail);

            return String.Format(CultureInfo.InvariantCulture,
            AuthorizeUrlNoResource,
            Uri.EscapeDataString(ClientId),
            Uri.EscapeDataString(RedirectUrl.ToString()),
            Uri.EscapeDataString(stateGuid));
        }

        [HttpPost()]
        public string CompleteOAuthFlow(AuthorizationParameters parameters) 
        {
            // Look up the email from the guid/user map.
            string userEmail = Storage.AppConfigCache.GetUserFromStateGuid(parameters.State);
            if (string.IsNullOrEmpty(userEmail))
            {
                // Per the Azure docs, the response from the auth code request has
                // to include the value of the state parameter passed in the request.
                // If it is not the same, then you should not accept the response.
                throw new HttpResponseException(Request.CreateErrorResponse(HttpStatusCode.OK, 
                    "Unknown state returned in OAuth flow."));
            }

            try
            {
                // Get authorized for the discovery service
                ClientCredential credential = new ClientCredential(ClientId, ClientSecret);
                string authority = string.Format(CultureInfo.InvariantCulture, OAuthUrl, "common");
                AuthenticationContext authContext = new AuthenticationContext(authority);
                AuthenticationResult result = authContext.AcquireTokenByAuthorizationCode(
                    parameters.Code, new Uri(RedirectUrl.GetLeftPart(UriPartial.Path)), credential, DiscoveryResource);

                // Cache the refresh token
                Storage.AppConfig appConfig = new Storage.AppConfig();
                appConfig.RefreshToken = result.RefreshToken;

                // Use the access token to get the user's OneDrive URL
                OneDriveServiceInfo serviceInfo = DiscoverServiceInfo(result.AccessToken);
                appConfig.OneDriveResourceId = serviceInfo.ResourceId;
                appConfig.OneDriveApiEndpoint = serviceInfo.Endpoint;

                // Save the user's configuration in our confic cache
                Storage.AppConfigCache.AddUserConfig(userEmail, appConfig);
                return "OAuth succeeded. Please close this window to continue.";
            }
            catch (ActiveDirectoryAuthenticationException ex)
            {
                throw new HttpResponseException(Request.CreateErrorResponse(HttpStatusCode.OK, 
                    "OAuth failed. " + ex.ToString()));
            }
        }

        public static OneDriveAccessDetails GetUsersOneDriveAccessDetails(string userEmail)
        {
            try
            {
                // Get the user's config, which contains the refresh token
                // and the OneDrive resource ID
                Storage.AppConfig appConfig = Storage.AppConfigCache.GetUserConfig(userEmail);

                // Request authorization for OneDrive
                ClientCredential credential = new ClientCredential(ClientId, ClientSecret);
                string authority = string.Format(CultureInfo.InvariantCulture, OAuthUrl, "common");
                AuthenticationContext authContext = new AuthenticationContext(authority);
                AuthenticationResult result = authContext.AcquireTokenByRefreshToken(
                    appConfig.RefreshToken, ClientId, credential, appConfig.OneDriveResourceId);

                // Update refresh token
                appConfig.RefreshToken = result.RefreshToken;
                Storage.AppConfigCache.AddUserConfig(userEmail, appConfig);

                return new OneDriveAccessDetails()
                {
                    ApiEndpoint = appConfig.OneDriveApiEndpoint,
                    AccessToken = result.AccessToken
                };
            }
            catch (ActiveDirectoryAuthenticationException)
            {
                return null;
            }
        }

        public static OneDriveServiceInfo DiscoverServiceInfo(string accessToken)
        {
            OneDriveServiceInfo serviceInfo = null;

            // Create a GET request to the discovery endpoint
            HttpWebRequest discoveryRequest =
                    WebRequest.CreateHttp(DiscoveryUrl);
            discoveryRequest.Headers.Add("Authorization", string.Format("Bearer {0}", accessToken));
            discoveryRequest.Method = "GET";
            discoveryRequest.Accept = "application/json";

            HttpWebResponse discoveryResponse = (HttpWebResponse)discoveryRequest.GetResponse();

            if (discoveryResponse.StatusCode == HttpStatusCode.OK)
            {
                Stream discoveryResponseStream = discoveryResponse.GetResponseStream();
                StreamReader discoveryReader = new StreamReader(discoveryResponseStream);
                string discoveryPayload = discoveryReader.ReadToEnd();

                // Deserialize the JSON response
                var discoveryResult = JsonConvert.DeserializeObject<dynamic>(discoveryPayload);

                foreach(var service in discoveryResult.value)
                {
                    // Look for the entry that matches OneDrive
                    if (String.Compare(service.capability.ToString(), OneDriveCapability) == 0)
                    {
                        serviceInfo = new OneDriveServiceInfo() 
                        { 
                            ResourceId = service.serviceResourceId.ToString(), 
                            Endpoint = service.serviceEndpointUri.ToString() 
                        };
                        break;
                    }
                }
            }

            return serviceInfo;
        }

        #region Helper classes
        public class AuthorizationRequest
        {
            public string UserEmail { get; set; }
        }

        public class AuthorizationParameters
        {
            public string Code { get; set; }
            public string State { get; set; }
        }

        public class OneDriveServiceInfo
        {
            public string ResourceId { get; set; }
            public string Endpoint { get; set; }
        }

        public class OneDriveAccessDetails
        {
            public string ApiEndpoint { get; set; }
            public string AccessToken { get; set; }
        }
        #endregion
    }
}

// MIT License: 

// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 

// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 