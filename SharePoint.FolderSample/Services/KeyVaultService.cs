using Microsoft.Azure.KeyVault;
using Microsoft.Azure.KeyVault.Models;
using Newtonsoft.Json.Linq;
using SharePoint.FolderSample.Code;
using SharePoint.FolderSample.Models;
using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace SharePoint.FolderSample.Services
{
    internal class KeyVaultService : IDisposable
    {
        readonly HttpClient client;
        readonly KeyVaultClient keyVault;

        public KeyVaultService()
        {
            client = new HttpClient();
            keyVault = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(GetTokenAsync), client);
        }

        public void Dispose()
        {
            keyVault.Dispose();
            client.Dispose();
        }

        public async Task<string> GetSecretAsync(string key)
        {
            SecretBundle secret = await keyVault.GetSecretAsync(Config.KeyVaultUrl, key);
            return secret.Value;
        }

        private async Task<string> GetTokenAsync(string authority, string resource, string scope)
        {
            JObject tokenResult = await AuthenticationUtilities.GetADAppTokenAsync(authority, resource, Config.ClientId, Config.ClientSecret);
            return tokenResult["access_token"].ToString();
        }
    }
}
