using System;
using System.Collections.Generic;

namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("WELCOME TO SILVERLIGHT RESEARCH MAIL SENDER");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Support : tech@silverlightresearch.com");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("=======================================================\n\n");
            Console.ForegroundColor = ConsoleColor.Gray;

            string dropboxUrl = Manager.AskUrl();
            DropboxTools dropboxtool = new DropboxTools(dropboxUrl);

            int templateChoice = Manager.AskTemplate();
            string dynContent = Manager.AskDynContent(templateChoice); //only for template 1
            string customBody = Manager.AskCustomBody(templateChoice); //only for template 3
            string mailSubject = Manager.AskSubject(templateChoice);
            List<string> regionCode = Manager.AskRegion();

            var dlTask = DropboxTools.Download(DropboxTools.m_dropboxFilePath);
            dlTask.Wait();

            ExcelReader Reader = new ExcelReader();
            Reader.Init("C:\\Temp\\DropboxDownloads\\" + DropboxTools.m_fileName, regionCode);

            Manager Manage = new Manager();
            Manage.SendMailFromExcelFile(templateChoice, mailSubject, dynContent, customBody, Reader);

            var upTask = DropboxTools.Upload(DropboxTools.m_dropboxFolderPath, DropboxTools.m_fileName);
            upTask.Wait();

            Console.WriteLine("Press any key to leave");
            Console.ReadKey();
        }
    }
}
