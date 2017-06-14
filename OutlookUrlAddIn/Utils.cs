using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace OutlookUrlAddIn
{
    public class Utils
    {
        /// <summary>
        /// Return url value if found.
        /// </summary>
        /// <param name="body"></param>
        /// <see cref="http://regexr.com/38l0t">Regexr</see>
        /// <returns></returns>
        public static string[] ExtractUrl(string body)
        {
            List<string> urls = new List<string>();
            string url = string.Empty;
            string pattern = @"(?:(?:https?:\/\/)|(?:www\.))[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,4}\b(?:[-a-zA-Z0-9@:%_\+.~#?&/=]*)";
            Regex regex = new Regex(pattern);
            MatchCollection collection = regex.Matches(body);
            foreach (Match match in collection)
            {
                if (match.Success)
                {
                    foreach (Group group in match.Groups)
                    {
                        urls.Add(group.Value);
                    }
                }
            }

            return urls.ToArray();
        }
    }
}