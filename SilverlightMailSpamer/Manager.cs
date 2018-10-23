using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailSender
{
    enum RegionID
    {
        Europe = 1,
        US = 2,
        Asia = 3,
        Worldwide = 4,
    }

    class Manager
    {
        public static string m_folderPath { get; set; }
        public static int m_mailsNumber;

        //send the mails from an array experts created by the "createArrayExperts" method in ExcelReader
        public void SendMailFromExcelFile(int templateChoice, string subject, string dynContent, string customBody, ExcelReader Reader)
        {
            for (int ListID = 0; ListID < Reader.GetExpertsArray().Count; ListID++)
            {
                if (Reader.GetExpertsArray()[ListID][2] != null)
                {
                    if (Reader.GetExpertsArray()[ListID][2].Contains("@"))
                    {
                        List<string> mails = new List<string>();
                        mails = OutlookTools.SplitMails(Reader.GetExpertsArray()[ListID][2]);

                        foreach (string mail in mails)
                        {
                            System.Threading.Thread.Sleep(1000);
                            m_mailsNumber = mails.Count();
                            if (templateChoice == 1)
                            {
                                if (!OutlookTools.CreateMailTemplate1(Convert.ToInt32(Reader.GetExpertsArray()[ListID][4]), Reader.GetExpertsArray()[ListID][7], Reader.GetExpertsArray()[ListID][0], Reader.GetExpertsArray()[ListID][1], mail, dynContent, subject))
                                {
                                    m_mailsNumber = m_mailsNumber - 1;
                                }
                                else
                                {
                                    Reader.WriteStatut("C:\\Temp\\DropboxDownloads\\" + DropboxTools.m_fileName, System.Convert.ToInt32(Reader.GetExpertsArray()[ListID][5]), 1);
                                }
                            }
                            else if (templateChoice == 2)
                            {   if (Reader.GetExpertsArray()[ListID][6] != "")
                                {
                                    if (!OutlookTools.CreateMailTemplate2(Convert.ToInt32(Reader.GetExpertsArray()[ListID][4]), Reader.GetExpertsArray()[ListID][7], Reader.GetExpertsArray()[ListID][0], Reader.GetExpertsArray()[ListID][1], mail, Reader.GetExpertsArray()[ListID][6], subject))
                                    {
                                        m_mailsNumber = m_mailsNumber - 1;
                                    }
                                    else
                                    {
                                        Reader.WriteStatut("C:\\Temp\\DropboxDownloads\\" + DropboxTools.m_fileName, System.Convert.ToInt32(Reader.GetExpertsArray()[ListID][5]), 1);
                                    }
                                }
                                else
                                {
                                    m_mailsNumber = 0;
                                }
                            }
                            else //template 3 (custom)
                            {
                                if (customBody.Contains("[companyname]"))
                                {
                                    if (Reader.GetExpertsArray()[ListID][6] != "")
                                    {
                                        if (!OutlookTools.CreateMailTemplate3(customBody, Convert.ToInt32(Reader.GetExpertsArray()[ListID][4]), Reader.GetExpertsArray()[ListID][7], Reader.GetExpertsArray()[ListID][0], Reader.GetExpertsArray()[ListID][1], mail, Reader.GetExpertsArray()[ListID][6], subject))
                                        {
                                            m_mailsNumber = m_mailsNumber - 1;
                                        }
                                        else
                                        {
                                            Reader.WriteStatut("C:\\Temp\\DropboxDownloads\\" + DropboxTools.m_fileName, System.Convert.ToInt32(Reader.GetExpertsArray()[ListID][5]), 1);
                                        }
                                    }
                                    else
                                    {
                                        m_mailsNumber = 0;
                                    }
                                }
                                else
                                {
                                    if (!OutlookTools.CreateMailTemplate3(customBody, Convert.ToInt32(Reader.GetExpertsArray()[ListID][4]), Reader.GetExpertsArray()[ListID][7], Reader.GetExpertsArray()[ListID][0], Reader.GetExpertsArray()[ListID][1], mail, Reader.GetExpertsArray()[ListID][6], subject))
                                    {
                                        m_mailsNumber = m_mailsNumber - 1;
                                    }
                                    else
                                    {
                                        Reader.WriteStatut("C:\\Temp\\DropboxDownloads\\" + DropboxTools.m_fileName, System.Convert.ToInt32(Reader.GetExpertsArray()[ListID][5]), 1);
                                    }
                                }
                            }
                        }
                        Console.WriteLine(m_mailsNumber + " mail(s) sent to: " + Reader.GetExpertsArray()[ListID][0] + "\n");
                    }
                }
                else
                {
                    Console.WriteLine("Unable to send mail to: " + Reader.GetExpertsArray()[ListID][0]);
                    Console.WriteLine("mails on Wrong column or missing mail adress?\n");
                }
            }
        }

        public static string AskDynContent(int templateChoice)
        {
            if (templateChoice == 1)
            {
                return AskTopic();
            }
            else
            {
                return string.Empty; //no dynamic content for the template 2 & 3
            }
        }

        public static string AskCustomBody(int templateChoice)
        {
            if (templateChoice == 3)
            {
                bool confirmed = false;
                string customBodyFilePath = string.Empty;
                string customBody = string.Empty;

                do
                {
                    Console.WriteLine("Select the custom body file path :");
                    var t = new Thread((ThreadStart)(() => {
                        OpenFileDialog browser = new OpenFileDialog
                        {
                            Title = "Select the custom body file path",
                            Filter = "Text files (*.txt) | *.txt",
                            FilterIndex = 1,
                            Multiselect = false
                        };
                        if (browser.ShowDialog() == DialogResult.OK)
                        {
                            customBodyFilePath = browser.FileName;
                        }
                    }));

                    t.SetApartmentState(ApartmentState.STA);
                    t.Start();
                    t.Join();
                    Console.WriteLine(customBodyFilePath);

                    ConsoleKey response;
                    do
                    {
                        Console.Write("Are you sure you want to validate this file? [y/n] ");
                        response = Console.ReadKey(false).Key;
                        if (response != ConsoleKey.Enter)
                            Console.WriteLine();

                    } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                    confirmed = response == ConsoleKey.Y;
                } while (!confirmed);

                var lines = File.ReadLines(customBodyFilePath);
                foreach (var line in lines)
                {
                    customBody = customBody + line + "<br>";
                }
                FileInfo FileInfo = new FileInfo(customBodyFilePath);
                m_folderPath = FileInfo.DirectoryName + "\\";
                return customBody;
            }
            else
            {
                return string.Empty; //no dynamic content for the template 2 & 3
            }
        }

        public static string AskTopic()
        {
            bool confirmedTopic = false;
            string topic;
            do
            {
                Console.WriteLine("\nType the project's topic and press enter :");
                Console.WriteLine("     - without any capital letter");
                Console.WriteLine("     - the sentance is : for a short call 30 minutes-1 hour maximum to talk about (topic)");
                Console.WriteLine("     - example : the oil transportation market");
                topic = Console.ReadLine();
                Console.WriteLine("\nYou entered, " + topic + " as topic");

                ConsoleKey response;
                do
                {
                    Console.Write("Are you sure you want to choose this topic to send mails? [y/n] ");
                    response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                    if (response != ConsoleKey.Enter)
                        Console.WriteLine();

                } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                confirmedTopic = response == ConsoleKey.Y;
            } while (!confirmedTopic);
            Console.WriteLine("\nTopic \"{0}\" confirmed", topic);

            return topic;
        }

        public static List<string> AskRegion()
        {
            bool confirmedRegion = false;
            List<string> regionList = new List<string>();
            do
            {
                bool regionFlag1 = false;
                bool regionFlag2 = false;

                do
                {
                    Console.WriteLine("\nChoose a region and press enter:");
                    Console.WriteLine("    1 - Europe");
                    Console.WriteLine("    2 - US");
                    Console.WriteLine("    3 - Asia");
                    Console.WriteLine("    4 - Worldwide");
                    regionList.Add(Console.ReadLine());
                    if (regionList[0] == "1" || regionList[0] == "2" || regionList[0] == "3" || regionList[0] == "4")
                    {
                        regionFlag1 = true;
                        Console.WriteLine("\nYou entered " + Enum.GetName(typeof(RegionID), System.Convert.ToInt32(regionList[0])) + " as region");
                    }

                    else
                    {
                        regionList.RemoveAt(0);
                        Console.WriteLine("You entered a wrong region number");
                    }
                } while (regionFlag1 == false);

                do
                {
                    if (regionList[0] != "4")
                    {
                        Console.WriteLine("Would you like to add another one? (Put the region number or just press enter if you don't)");
                        regionList.Add(Console.ReadLine());
                        if (regionList[1] == "")
                        {
                            regionFlag2 = true;
                        }
                        if (regionList[0] == regionList[1])
                        {
                            regionList.RemoveAt(1);
                            Console.WriteLine("You already entered this region");
                        }
                        else
                        {
                            if (regionList[1] == "1" || regionList[1] == "2" || regionList[1] == "3" || regionList[1] == "")
                            {
                                regionFlag2 = true;
                            }
                            else
                            {
                                regionList.RemoveAt(1);
                                Console.WriteLine("You entered a wrong region number");
                            }
                        }
                    }
                    else
                    {
                        regionFlag2 = true;
                        regionList.Add("");
                    }
                } while (regionFlag2 == false);


                ConsoleKey response;

                do
                {
                    if (regionList[1] == "")
                    {
                        Console.Write("Are you sure you want to choose " + Enum.GetName(typeof(RegionID), System.Convert.ToInt32(regionList[0])) + " ?[y/n] ");
                        regionList.RemoveAt(1);
                    }
                    else
                    {
                        Console.Write("Are you sure you want to choose " + Enum.GetName(typeof(RegionID), System.Convert.ToInt32(regionList[0])) + " and " + Enum.GetName(typeof(RegionID), System.Convert.ToInt32(regionList[1])) + " ?[y/n] ");
                    }

                    response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                    if (response != ConsoleKey.Enter)
                        Console.WriteLine();
                    if (response == ConsoleKey.N)
                        regionList.Clear();

                } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                confirmedRegion = response == ConsoleKey.Y;
            } while (!confirmedRegion);

            return regionList;
        }

        public static string AskSubject(int templateChoice)
        {
            bool confirmedTopic = false;
            string defaultSubject;
            string subject;

            if (templateChoice == 1 || templateChoice == 3)
            {
                defaultSubject = "Call with Silverlight";
            }
            else
            {
                defaultSubject = "Call with Silverlight Research Expert Network";
            }

            do
            {
                Console.WriteLine("\nType the subject's mail and press enter :");
                Console.WriteLine("     - by default : "+ defaultSubject + " (just press enter)");

                subject = Console.ReadLine();
                if (subject == "")
                    subject = defaultSubject;
                Console.WriteLine("\nYou entered, " + subject + " as subject");

                ConsoleKey response;
                do
                {
                    Console.Write("Are you sure you want to choose this subject to send mails? [y/n] ");
                    response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                    if (response != ConsoleKey.Enter)
                        Console.WriteLine();

                } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                confirmedTopic = response == ConsoleKey.Y;
            } while (!confirmedTopic);
            Console.WriteLine("\nTopic \"{0}\" confirmed", subject);

            return subject;
        }

        public static string AskCompanyName()
        {
            bool confirmedCompany = false;
            string company;
            do
            {
                Console.WriteLine("\nType the company name of the client and press enter :");
                Console.WriteLine("     - the sentance is : to arrange a short call to learn more about your work at  (company name)");
                Console.WriteLine("     - example : LEK Consulting");
                company = Console.ReadLine();
                Console.WriteLine("\nYou entered, " + company + " as company name");

                ConsoleKey response;
                do
                {
                    Console.Write("Are you sure you want to choose this company name to send mails? [y/n] ");
                    response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                    if (response != ConsoleKey.Enter)
                        Console.WriteLine();

                } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                confirmedCompany = response == ConsoleKey.Y;
            } while (!confirmedCompany);
            Console.WriteLine("Company \"{0}\" confirmed", company);

            return company;
        }

        public static int AskTemplate()
        {
            bool confirmedTemplate = false;
            int template;
            string sTemplate;

            do
            {
                Console.WriteLine("\n");
                Console.Write("     - Template 1 : Expert Template\nGreetings,\nI hope you are well.\nI am writing to you from Silverlight Research, a consultancy firm based in central London. We are currently working on a project in your field of expertise and would be keen to speak for a 5 minute call today if possible.\nWe are working with a client who would be very interested to be introduced to you for a short call 30 minutes-1 hour maximum to talk about ");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("[TOPIC]");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write(".\nGiven your experience in this field we would be really keen to organise a short introduction call.\nPlease could you let me know the number to reach you and let me know what time works to discuss this briefly.\nMany thanks, I look forward to speaking with you.\nKind regards,\nSignature\n\n");
                Console.Write("     - Template 2 : Clients Template\nGreetings,\nI hope you are well.\nI wanted to reach out and would be very interested to arrange a short call to learn more about your work at ");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("[COMPANY NAME]");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write(" as well as introduce what we do at Silverlight Research.\nWe are a Specialist Knowledge Research house and Expert Network based in London. Silverlight research is run by ex-investment bankers from Morgan Stanley and our value add versus other providers is a highly detailed and bespoke yet cost efficient service.\nWe work primarily with large investment and consulting firms, but do not currently work with your team.\nPlease could we arrange a call? We would be very interested in exploring how our services may be useful for you. I look forward to speaking.\nKind regards,\nSignature\n\n");
                Console.Write("     - Template 3 : Custom Template\n\n");

                do
                {
                    Console.WriteLine("\nType number of the template that you want to use to mail, and press enter :");
                    sTemplate = Console.ReadLine();
                }
                while (sTemplate != "1" && sTemplate != "2" && sTemplate != "3");

                template = Convert.ToInt32(sTemplate);
                Console.WriteLine("\nYou entered, " + template + " as template choice");

                ConsoleKey response;
                do
                {
                    Console.Write("Are you sure you want to choose this template to send mails? [y/n] ");
                    response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                    if (response != ConsoleKey.Enter)
                        Console.WriteLine();

                } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                confirmedTemplate = response == ConsoleKey.Y;
            } while (!confirmedTemplate);
            Console.WriteLine("\nTemplate \"{0}\" confirmed", template);

            return template;
        }

        public static string AskUrl()
        {
            bool confirmedUrl = false;
            string url;
            do
            {
                do
                {
                    Console.WriteLine("Paste the URL of the Dropbox Excel file and press enter :");
                    url = Console.ReadLine();
                }
                while (!url.Contains("/") && !url.Contains("\\"));

                ConsoleKey response;
                do
                {
                    Console.Write("Do you want to confirm this file URL to send mails? [y/n] ");
                    response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                    if (response != ConsoleKey.Enter)
                        Console.WriteLine();

                } while (response != ConsoleKey.Y && response != ConsoleKey.N);

                confirmedUrl = response == ConsoleKey.Y;
            } while (!confirmedUrl);
            Console.WriteLine("\nFile URL confirmed");

            return url;
        }
    }
}
