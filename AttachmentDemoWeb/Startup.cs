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
                ValidAudience = "05adb30e-50fa-4ae2-9cec-eab2cd6095b0",
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
