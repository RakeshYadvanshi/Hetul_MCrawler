using System;

namespace Shared
{
    public class Helper
    {
        public static HtmlAgilityPack.HtmlDocument GetContentFromUrl(string url)
        {
            HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb
            {
                UserAgent =
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
            };
            HtmlAgilityPack.HtmlDocument doc = web.Load(url);
            return doc;
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
       
    }
}
