using System;
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
            return str.ToLower().Replace("days ago", "").Replace("day ago", "").Replace("just posted", "0").Replace("today", "0");
        }

    }
}
