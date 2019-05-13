using System.Configuration;

namespace SharePoint.FolderSample.Models
{
    internal static class Config
    {
        public static string SiteUrl { get { return ConfigurationManager.AppSettings.Get(nameof(SiteUrl)); } }
        public static string ClientId { get { return ConfigurationManager.AppSettings.Get(nameof(ClientId)); } }
        public static string ClientSecret { get { return ConfigurationManager.AppSettings.Get(nameof(ClientSecret)); } }
        public static string KeyVaultUrl { get { return ConfigurationManager.AppSettings.Get(nameof(KeyVaultUrl)); } }

        public static bool IsValid
        {
            get
            {
                return !(
                    string.IsNullOrWhiteSpace(SiteUrl) ||
                    string.IsNullOrWhiteSpace(ClientId) ||
                    string.IsNullOrWhiteSpace(ClientSecret) ||
                    string.IsNullOrWhiteSpace(KeyVaultUrl)
                );
            }
        }
    }
}
