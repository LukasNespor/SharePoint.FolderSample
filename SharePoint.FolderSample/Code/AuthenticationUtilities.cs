using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SharePoint.FolderSample.Services;
using System;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Web;

namespace SharePoint.FolderSample.Code
{
    internal static class AuthenticationUtilities
    {
        public async static Task<ClientContext> GetClientContextAsync(string siteUrl)
        {
            string tenantName = new Uri(siteUrl).Host.Replace(".sharepoint.com", "");
            string authority = $"https://login.microsoftonline.com/{tenantName}.onmicrosoft.com/oaut/token2";
            string resource = $"https://{tenantName}.sharepoint.com/";

            string clientId = string.Empty;
            X509Certificate2 certificate = null;
            using (KeyVaultService keyVault = new KeyVaultService())
            {
                clientId = await keyVault.GetSecretAsync(Constants.SharePointClientId);
                var certPfx = await keyVault.GetSecretAsync(Constants.SharePointCert);
                certificate = new X509Certificate2(
                    Convert.FromBase64String(certPfx),
                    (string)null,
                    X509KeyStorageFlags.MachineKeySet);
            }

            var authenticationContext = new AuthenticationContext(authority, false);
            AuthenticationResult authResult = await authenticationContext.AcquireTokenAsync(
                resource,
                new ClientAssertionCertificate(clientId, certificate));

            var ctx = new ClientContext(siteUrl);
            ctx.ExecutingWebRequest += (s, e) =>
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = $"Bearer {authResult.AccessToken}";
            };

            return ctx;
        }

        public static async Task<JObject> GetADAppTokenAsync(string authority, string audience, string clientId, string clientSecret)
        {
            string loginUrl = string.Format("{0}/oauth2/token", authority);

            WebRequest request = WebRequest.Create(loginUrl);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";

            string content = string.Format(
                "resource={0}&client_id={1}&client_secret={2}&grant_type=client_credentials",
                HttpUtility.UrlEncode(audience),
                HttpUtility.UrlEncode(clientId),
                HttpUtility.UrlEncode(clientSecret));

            return await GetResponseAsync(request, content);
        }

        private static async Task<JObject> GetResponseAsync(WebRequest request, string content)
        {
            using (StreamWriter writer = new StreamWriter(request.GetRequestStream()))
            {
                writer.Write(content);
            }

            try
            {
                WebResponse response = await request.GetResponseAsync();
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string responseContent = reader.ReadToEnd();
                    return JsonConvert.DeserializeObject<JObject>(responseContent);
                }
            }
            catch (WebException webException)
            {
                if (webException.Response != null)
                {
                    using (StreamReader reader = new StreamReader(webException.Response.GetResponseStream()))
                    {
                        string responseContent = reader.ReadToEnd();
                    }
                }
            }

            return null;
        }
    }
}
