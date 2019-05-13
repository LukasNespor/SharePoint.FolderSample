using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using SharePoint.FolderSample.Code;
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
                Console.WriteLine("Missing argument with folder URL or configuration values in app settings.");
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            string folder = args[0];

            using (ClientContext client = await AuthenticationUtilities.GetClientContextAsync(Config.SiteUrl))
            using (SecureString password = new SecureString())
            {
                Console.WriteLine("Connecting to SharePoint site...");
                client.Load(client.Web, w => w.ServerRelativeUrl);
                await client.ExecuteQueryAsync();

                Console.WriteLine("Loading files list...");
                var files = await GetFilesFromFolderAsync(client, folder);
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
        /// Query SharePoint Search and fetch information about each file (URL and name).
        /// </summary>
        /// <param name="client">Existing SharePoint ClientContext</param>
        /// <param name="folderUrl">Full server folder URL</param>
        /// <returns>List of files information</returns>
        static async Task<IEnumerable<SPFileInfo>> GetFilesFromFolderAsync(ClientContext client, string folderUrl)
        {
            KeywordQuery query = new KeywordQuery(client)
            {
                QueryText = $"IsContainer=false Path:\"{folderUrl}\"",
                RowLimit = 500
            };
            SearchExecutor searchExecutor = new SearchExecutor(client);
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(query);
            await client.ExecuteQueryAsync();

            var relevantResults = results.Value.Where(x => x.TableType == "RelevantResults").FirstOrDefault();
            if (relevantResults == null)
                return Array.Empty<SPFileInfo>();

            var files = new List<SPFileInfo>();
            foreach (IDictionary<string, object> item in relevantResults.ResultRows)
            {
                var fileKey = item.Where(x => x.Key == "Path").FirstOrDefault();
                Uri filePath = new Uri(fileKey.Value.ToString());

                files.Add(new SPFileInfo()
                {
                    ServerRelativeUrl = filePath.PathAndQuery,
                    FileName = filePath.PathAndQuery.Substring(filePath.PathAndQuery.LastIndexOf("/") + 1)
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
            return $"{baseUrl.Trim('/')}/{string.Join("/", cleaned)}";
        }

        static bool IsConfigAndArgumentsValid(string[] args)
        {
            return args.Length > 0 && Config.IsValid;
        }
    }
}
