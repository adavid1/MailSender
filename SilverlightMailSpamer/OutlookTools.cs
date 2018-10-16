using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailSender
{
    class OutlookTools
    {
        public static bool CreateMailTemplate1(int greetings, string gender, string recipientFirstname, string recipientLastname, string recipientMail, string topic, string mailSubject)
        {
            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = recipientMail;
                mail.Subject = mailSubject;
                mail.HTMLBody = String.Format("<html><body style='font-size:10.0pt'><p><font face=\"calibri\"> " + Greetings.GetGreetingsEN(greetings, gender, recipientFirstname, recipientLastname) + "," +
                                              "<br> I hope you are well. </p>" +
                                              "<p> I am writing to you from Silverlight Research, a consultancy firm based in central London. We are currently working on a project in your field of expertise and would be keen to speak for a 5 minute call today if possible. We are working with a client who would be very interested to be introduced to you for a short call 30 minutes-1 hour maximum to talk about " + topic + ".  Given your experience in this field we would be really keen to organise a short introduction call.<p>" +
                                              "<p> Please could you let me know the number to reach you and let me know what time works to discuss this briefly. </p>" +
                                              "<p> Many thanks, I look forward to speaking with you. </p>" +
                                              "<p> Kind regards,  </p>" +
                                              "</font></body></html>") + ReadSignature();
                mail.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine("mail error : " + ex.Message);
                return false;
            }
            return true;
        }

        public static bool CreateMailTemplate2(int greetings, string gender, string recipientFirstname, string recipientLastname, string recipientMail, string companyName, string mailSubject)
        {
            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = recipientMail;
                mail.Subject = mailSubject;
                mail.HTMLBody = String.Format("<html><body style='font-size:10.0pt'><p><font face=\"calibri\"> " + Greetings.GetGreetingsEN(greetings, gender, recipientFirstname, recipientLastname) + "," +
                                              "<br> I hope you are well. </p>" +
                                              "<p> I wanted to reach out and would be very interested to arrange a short call to learn more about your work at " + companyName + " as well as introduce what we do at Silverlight Research. We are a Specialist Knowledge Research house and Expert Network based in London. Silverlight research is run by ex-investment bankers from Morgan Stanley and our value add versus other providers is a highly detailed and bespoke yet cost efficient service. We work primarily with large investment and consulting firms, but do not currently work with your team. <p>" +
                                              "<p> Please could we arrange a call? We would be very interested in exploring how our services may be useful for you. I look forward to speaking. </p>" +
                                              "<p> Kind regards,  </p>" +
                                              "</font></body></html>") + ReadSignature();
                mail.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine("mail error : " + ex.Message);
                return false;
            }
            return true;
        }

        public static bool CreateMailTemplate3(string customBody, int greetings, string gender, string recipientFirstname, string recipientLastname, string recipientMail, string companyName, string mailSubject)
        {
            try
            {
                if(customBody.Contains("[greetings-"))
                {
                    customBody = AddCustomGreetings(customBody, greetings, gender, recipientFirstname, recipientLastname);
                }
                else
                {
                    Console.WriteLine("Error : Wrong greetings format");
                    Console.ReadKey();
                    Environment.Exit(-1);
                }

                customBody = customBody.Replace("[firstname]", recipientFirstname);
                customBody = customBody.Replace("[lastname]", recipientLastname);
                customBody = customBody.Replace("[companyname]", companyName);

                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = recipientMail;
                mail.Subject = mailSubject;
                mail.HTMLBody = String.Format("<html><body style='font-size:10.0pt'>" + customBody + "<br></font></body></html>") + ReadSignature();
                mail.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine("mail error : " + ex.Message);
                return false;
            }
            return true;
        }

        public static string ReadSignature()
        {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string secondAppDataDir = @"C:\Users\"+ Environment.UserName+ @"\AppData\Local\Packages\Microsoft.Office.Desktop_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Signatures";
            string signature = string.Empty, content = string.Empty;

            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

                if (fiSignature.Length > 0)
                {
                    foreach (var selectedSignature in fiSignature)
                    {
                        StreamReader sr = new StreamReader(selectedSignature.FullName, Encoding.Default);
                        content = sr.ReadToEnd();

                        if (content.Contains("@silverlightresearch.com"))
                        {
                            signature = content;
                        }
                    }
                }
            }

            diInfo = new DirectoryInfo(secondAppDataDir);
            content = string.Empty;

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature2 = diInfo.GetFiles("*.htm");

                if (fiSignature2.Length > 0)
                {
                    foreach (var selectedSignature in fiSignature2)
                    {
                        StreamReader sr2 = new StreamReader(selectedSignature.FullName, Encoding.Default);
                        content = sr2.ReadToEnd();

                        if (content.Contains("silverlightresearch"))
                        {
                            signature = content;
                        }
                    }
                }
            }
            return signature;
        }

        public static List<string> SplitMails(string scope)
        {
            List<string> splitedMails = new List<string>();

            var mails = scope.Split(new char[] { ' ', '\n', ';' });

            foreach (string mail in mails)
            {
                if (mail != "")
                    splitedMails.Add(mail);
            }

            return splitedMails;
        }

        public static string AddCustomGreetings(string body, int greetingsKey, string genderKey, string firstname, string lastname)
        {
            if (body.Contains("[greetings-en]"))
            {
                return body.Replace("[greetings-en]", Greetings.GetGreetingsEN(greetingsKey, genderKey, firstname, lastname));
            }
            else if (body.Contains("[greetings-de]"))
            {
                return body.Replace("[greetings-de]", Greetings.GetGreetingsDE(greetingsKey, genderKey, firstname, lastname));
            }
            else if (body.Contains("[greetings-fr]"))
            {
                return body.Replace("[greetings-fr]", Greetings.GetGreetingsFR(greetingsKey, genderKey, firstname, lastname));
            }
            else
            {
                Console.WriteLine("Error : Wrong greetings format, leave the app");
                Console.ReadKey();
                return string.Empty;
            }
        }
    }
}
