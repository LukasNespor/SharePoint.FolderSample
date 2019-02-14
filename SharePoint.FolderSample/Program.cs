using Microsoft.SharePoint.Client;
using SharePoint.FolderSample.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Threading.Tasks;
using IO = System.IO;

namespace SharePoint.FolderSample
{
    class Program
    {
        /// <summary>
        /// Load list of files in SharePoint folder and then downloads each file to 'Files' folder nex to the exe.
        /// Example of first argument: folder/another subfolder/and so on
        /// </summary>
        /// <param name="args">List relative URL to the folder</param>
        /// <returns>Task</returns>
        static async Task Main(string[] args)
        {
            if (!IsConfigAndArgumentsValid(args))
            {
                Console.WriteLine("Missing argument with folder URL of configuration values in app settings.");
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            string folder = args[0];

            using (ClientContext client = new ClientContext(Config.SiteUrl))
            using (SecureString password = new SecureString())
            {
                Console.WriteLine("Connecting to SharePoint site...");
                Config.Password.ToList().ForEach(c => password.AppendChar(c));
                client.Credentials = new SharePointOnlineCredentials(Config.Username, password);
                client.Load(client.Web, w => w.ServerRelativeUrl);
                await client.ExecuteQueryAsync();

                Console.WriteLine("Loading files list...");
                var files = await GetFilesFromFolderAsync(client, Config.WebRelativeListUrl, folder);
                Console.WriteLine($"Loaded {files.Count()} files");
                IO.Directory.CreateDirectory("Files");

                foreach (SPFileInfo fileInfo in files)
                {
                    Console.Write($"Downloading file '{fileInfo.FileName}' ... ");
                    using (IO.Stream stream = await DownloadFileAsync(client, fileInfo.ServerRelativeUrl))
                    using (IO.FileStream file = IO.File.Create($@"Files\{fileInfo.FileName}"))
                    {
                        await stream.CopyToAsync(file);
                        Console.WriteLine("OK");
                    }
                }
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        /// <summary>
        /// Query SharePoint list (library) and fetch information about each file (URL and name).
        /// </summary>
        /// <param name="client">Existing SharePoint ClientContext</param>
        /// <param name="webRelativeListUrl">Web relative list URL. Example: 'Shared Documents'</param>
        /// <param name="listRelativeFolderUrl">List relative folder URL. Example: 'folder/subfolder'</param>
        /// <returns>List of files information</returns>
        static async Task<IEnumerable<SPFileInfo>> GetFilesFromFolderAsync(
            ClientContext client, string webRelativeListUrl, string listRelativeFolderUrl)
        {
            List list = client.Web.GetList(CombineUrls(client.Web.ServerRelativeUrl, webRelativeListUrl));

            var query = new CamlQuery
            {
                FolderServerRelativeUrl = CombineUrls(client.Web.ServerRelativeUrl, webRelativeListUrl, listRelativeFolderUrl),
                ViewXml = "<View Scope='Recursive' />"
            };

            ListItemCollection items = list.GetItems(query);
            client.Load(items, i => i.Include(x => x.File.ServerRelativeUrl, x => x.File.Name));
            await client.ExecuteQueryAsync();

            var files = new List<SPFileInfo>();
            foreach (ListItem item in items)
            {
                files.Add(new SPFileInfo()
                {
                    ServerRelativeUrl = item.File.ServerRelativeUrl,
                    FileName = item.File.Name
                });
            }

            return files;
        }

        /// <summary>
        /// Download file from SharePoint as a Stream.
        /// </summary>
        /// <param name="client">Existing SharePoint ClientContext</param>
        /// <param name="serverRelativeFileUrl">Server relative URL to the file</param>
        /// <returns>Stream</returns>
        static async Task<IO.Stream> DownloadFileAsync(ClientContext client, string serverRelativeFileUrl)
        {
            File file = client.Web.GetFileByServerRelativeUrl(serverRelativeFileUrl);
            ClientResult<IO.Stream> result = file.OpenBinaryStream();
            await client.ExecuteQueryAsync();
            return result.Value;
        }

        /// <summary>
        /// Combine urls no matter the leading or trailing slash.
        /// </summary>
        /// <param name="baseUrl">Base part of result URL</param>
        /// <param name="otherParts">Other parts of result URL</param>
        /// <returns>Concatenated URL</returns>
        static string CombineUrls(string baseUrl, params string[] otherParts)
        {
            string[] cleaned = otherParts.ToList().Select(u => u.Trim('/')).ToArray();
            return $"/{baseUrl.Trim('/')}/{string.Join("/", cleaned)}";
        }

        static bool IsConfigAndArgumentsValid(string[] args)
        {
            return args.Length > 0 && Config.IsValid;
        }
    }
}
