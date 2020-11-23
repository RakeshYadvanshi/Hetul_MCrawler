using System;
using System.Threading;

namespace Shared
{
    public static class Helper
    {
        public static HtmlAgilityPack.HtmlDocument GetContentFromUrl(string url)
        {
            try
            {
                HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb
                {
                    UserAgent =
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
                };
                HtmlAgilityPack.HtmlDocument doc = web.Load(url);
                return doc;

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Thread.Sleep(5000);
                return GetContentFromUrl(url);
            }
            
         
        }

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
            return str.ToLower().Replace("days ago", "").Replace("day ago", "").Replace("just posted", "0").Replace("today", "0");
        }

    }
}
