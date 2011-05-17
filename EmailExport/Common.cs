using System;
using System.Collections.Generic;
using System.Text;

namespace EmailExport
{
    public class Common
    {
        public static List<string> ParseEmailAddresses(string raw)
        {
            List<string> temp = new List<string>();

            foreach(var email in raw.Split(';'))
                temp.Add(email);

            return temp;
        }

        public static bool ValidateEmail(string email)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(
                     email,
                     @"^(?("")("".+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +
                     @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$"
                     );
        }
    }
}
