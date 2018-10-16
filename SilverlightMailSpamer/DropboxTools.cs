using Dropbox.Api;
using Dropbox.Api.Files;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Web;

/// <summary>
/// Uses Dropbox.API to interact
/// 
/// To use in Program.cs :
///     var task = Dropbox.Function(parameters);
///     task.Wait();
/// </summary>

namespace MailSender
{
    class DropboxTools
    {
        private const string token = "";
        private const string basePath = "C:/Temp/DropboxDownloads/";

        public static string m_dropboxFilePath;
        public static string m_dropboxFolderPath;
        public static string m_fileName;

        public DropboxTools(string DropboxUrl)
        {
            m_dropboxFilePath = DropboxTools.GetFilePathFromUrl(DropboxUrl);
            m_dropboxFolderPath = DropboxTools.GetFolderPathFromUrl(DropboxUrl);
            m_fileName = Path.GetFileName(m_dropboxFilePath);
        }

        public static async Task ShowUsers()
        {
            using (var dbx = new DropboxClient(token))
            {
                var full = await dbx.Users.GetCurrentAccountAsync();
                Console.WriteLine("{0} - {1}", full.Name.DisplayName, full.Email);
                Console.ReadKey();
            }
        }

        public static async Task Download(string inputPath)
        {
            Console.WriteLine("\nDownload file...");

            if (!Directory.Exists(basePath))
            {
                Directory.CreateDirectory(basePath);
            }

            string outputPath = basePath + Path.GetFileName(inputPath);

            using (var dbx = new DropboxClient(token))
            {
                using (var response = await dbx.Files.DownloadAsync(inputPath))
                {
                    using (var fileStream = File.Create(outputPath))
                    {
                        (await response.GetContentAsStreamAsync()).CopyTo(fileStream);
                    }
                }
                Console.WriteLine("File downloaded");
            }
        }

        public static async Task Upload(string outputFolderPath, string inputFileName)
        {
            Console.WriteLine("\nUpload file...");

            string inputPath = basePath + inputFileName;
            string outputPath = outputFolderPath + "[emailed] " + inputFileName;

            using (var dbx = new DropboxClient(token))
            {
                using (var stream = File.OpenRead(inputPath))
                {
                    var response = await dbx.Files.UploadAsync(outputPath, WriteMode.Overwrite.Instance, body: stream);

                    Console.WriteLine("Uploaded Id {0} Rev {1}", response.Id, response.Rev);
                }
            }

            Console.WriteLine("File uploaded");
        }

        public static string GetFilePathFromUrl(string url)
        {
            StringBuilder builder = new StringBuilder(HttpUtility.UrlDecode(url));
            builder.Replace("https://www.dropbox.com/preview", "");
            builder.Replace("https://www.dropbox.com/home", "");
            builder.Replace("https://www.dropbox.com/ow/msft/edit/personal", "");
            builder.Replace("?preview=", "/");
            builder.Replace("?role=personal", "");

            return Convert.ToString(builder).Split('?')[0];
        }

        public static string GetFolderPathFromUrl(string url)
        {
            string newUrl = GetFilePathFromUrl(url);

            StringBuilder builder = new StringBuilder(newUrl);
            builder.Replace(Path.GetFileName(newUrl), "");

            return builder.ToString();
        }
    }
}
