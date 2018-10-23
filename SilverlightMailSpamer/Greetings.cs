using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MailSender
{
    class Greetings
    {
        public static string GetGreetingsDefault(int greetingsKey, string genderKey, string firstname, string lastname)
        {
            string sGreetings;

            if (greetingsKey == 1) //Europe & world (Dear M. lastname) 
            {
                if (genderKey == "1")
                {
                    sGreetings = "Dear Mr " + lastname;
                }
                else if (genderKey == "2")
                {
                    sGreetings = "Dear Ms " + lastname;
                }
                else
                {
                    sGreetings = "Dear M. " + lastname;
                }
            }
            else //UK & US (Dear firstname)
            {
                sGreetings = "Dear " + firstname;
            }

            return sGreetings;
        }

        public static string GetGreetingsCustom(int greetingsKey, string genderKey, string firstname, string lastname, string greetingLanguage)
        {
            string sGreetings = "";
            XmlDocument doc = new XmlDocument();
            doc.Load(Manager.m_folderPath + "GreetingsList.xml");
            
            foreach (XmlNode rootNode in doc.DocumentElement.ChildNodes)
            {
                if (rootNode.Attributes.GetNamedItem("name").Value.ToString() == greetingLanguage ) //or loop through its children as well
                {
                    foreach (XmlNode languageNode in rootNode.ChildNodes)
                    {
                        if (genderKey == "1" && languageNode.Name == "Male" || genderKey == "2" && languageNode.Name == "Female")
                        {
                            sGreetings = languageNode.InnerText;
                            sGreetings.TrimStart();
                            sGreetings.TrimEnd();
                            sGreetings = sGreetings + " " + lastname;
                            return sGreetings;
                        }
                        else
                        {
                            if (languageNode.Name == "Default")
                            {
                                sGreetings = languageNode.InnerText;
                                sGreetings.TrimStart();
                                sGreetings.TrimEnd();
                                sGreetings = sGreetings + " " + firstname;
                                return sGreetings;
                            }
                        }
                    }
                }
            }
            return "";
        }
    }
}
