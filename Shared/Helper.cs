using System;
using System.Linq;
using System.Threading;

namespace Shared
{
    public static class Helper
    {

        public static string GetAmazonJobId(string amazonHtmlContent)
        {
            var document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(amazonHtmlContent);
            var textContent = document.DocumentNode.QuerySelector(".first")?.InnerHtml;
            if (textContent != null)
            {
                textContent = textContent.Replace("Job ID:", "").Trim().Split('|')[0].Trim();
                return textContent;
            }
            return "";
        }

        public static string HandleEmptyUrl(this string url) => string.IsNullOrEmpty(url) ? "Application Form" : url;

        public static string HandleStringDateFromIndeed(this string str)
        {
             var jobWage=str.ToLower().Replace("days ago", "").Replace("day ago", "").Replace("just posted", "0").Replace("today", "0");

            while (jobWage.IndexOf('+') > -1)
            {
                jobWage = jobWage.Replace("+", "");
            }

            return jobWage.Trim();
        }
        public static string HandleJobWageFromIndeed(this string jobWage)
        {
            if (jobWage.Count(x => x == '$') > 1 && jobWage.IndexOf('-') > -1)
            {
                jobWage = jobWage.Split('-')[1];
                while (jobWage.IndexOf('+') > -1)
                {
                    jobWage = jobWage.Replace("+", "");
                }
            }

            return jobWage.ToLower().Replace("from","").Replace("an","").Replace("hour","")
                //.Replace("up to", "")
                //.Replace("a year", "")
                .Replace("++", "").Replace("up to","")
                .Replace("a week", "")
                .Trim();
        }
        public static string HandleStringJobLocationFromIndeed(this string jobLocation)
        {
            if (jobLocation.IndexOf(",", StringComparison.Ordinal) > -1)
            {
                var joblocationArray = jobLocation.Split(',');
                jobLocation = joblocationArray[0] + ", " + joblocationArray[1].Trim().Split(' ')[0];
            }

            return jobLocation;
        }


        public static string HtmlDecode(this string str)
        {
            return System.Net.WebUtility.HtmlDecode(str);
        }
    }
}
