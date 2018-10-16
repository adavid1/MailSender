using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSender
{
    class Greetings
    {
        public static string GetGreetingsEN(int greetingsKey, string genderKey, string firstname, string lastname)
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


        public static string GetGreetingsFR(int greetingsKey, string genderKey, string firstname, string lastname)
        {
            string sGreetings;

            if (greetingsKey == 1) //Europe & world (Dear M. lastname) 
            {
                if (genderKey == "1")
                {
                    sGreetings = "Cher Mr " + lastname;
                }
                else if (genderKey == "2")
                {
                    sGreetings = "Chère Mme " + lastname;
                }
                else
                {
                    sGreetings = "Cher M. " + lastname;
                }
            }
            else //UK & US (Dear firstname)
            {
                sGreetings = "Cher " + firstname;
            }

            return sGreetings;
        }


        public static string GetGreetingsDE(int greetingsKey, string genderKey, string firstname, string lastname)
        {
            string sGreetings;

            if (greetingsKey == 1) //Europe & world (Dear M. lastname) 
            {
                if (genderKey == "1")
                {
                    sGreetings = "Lieber Herr " + lastname;
                }
                else if (genderKey == "2")
                {
                    sGreetings = "Liebe Frau " + lastname;
                }
                else
                {
                    sGreetings = "Sehr geehrte Damen und Herren " + lastname;
                }
            }
            else //UK & US (Dear firstname)
            {
                sGreetings = "Lieber " + firstname;
            }

            return sGreetings;
        }
    }
}
