using AttachmentDemoWeb.App_Start;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Owin;
using Microsoft.Owin.Security.Jwt;
using Microsoft.Owin.Security.OAuth;
using Owin;
using System.Configuration;

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
                ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
                // Microsoft Accounts have an issuer GUID that is different from any organizational tenant GUID,
                // so to support both kinds of accounts, we do not validate the issuer.
                ValidateIssuer = false,
                SaveSigninToken = true
            };

            string[] endAuthoritySegments = { "oauth2/v2.0" };
            string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
            string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

            app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            {
                AccessTokenFormat = new JwtFormat(tokenValidationParms, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
            });
        }
    }
}
