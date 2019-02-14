using System.Configuration;

namespace SharePoint.FolderSample.Models
{
    internal static class Config
    {
        public static string SiteUrl { get { return ConfigurationManager.AppSettings.Get(nameof(SiteUrl)); } }
        public static string WebRelativeListUrl { get { return ConfigurationManager.AppSettings.Get(nameof(WebRelativeListUrl)); } }
        public static string Username { get { return ConfigurationManager.AppSettings.Get(nameof(Username)); } }
        public static string Password { get { return ConfigurationManager.AppSettings.Get(nameof(Password)); } }
        public static bool IsValid
        {
            get
            {
                return !(
                    string.IsNullOrWhiteSpace(SiteUrl) ||
                    string.IsNullOrWhiteSpace(WebRelativeListUrl) ||
                    string.IsNullOrWhiteSpace(Username) ||
                    string.IsNullOrWhiteSpace(Password)
                );
            }
        }
    }
}
