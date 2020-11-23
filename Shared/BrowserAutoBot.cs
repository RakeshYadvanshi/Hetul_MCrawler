using System;
using System.Linq;
using System.Threading.Tasks;
using PuppeteerSharp;

namespace Shared
{
    public class BrowserAutoBot
    {
        private static Browser _browser;
        public static async Task<Page> setupBrowser()
        {
            _browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
                Headless = false,
                ExecutablePath =
                      @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", //"C:\Users\Admin\AppData\Local\Chromium\Application\chrome.exe",
                LogProcess = true,
                IgnoreHTTPSErrors = true,
                Args = new[]
                  {
                    "--no-sandbox",
                    "--incognito",
                    "--disable-infobars",
                    "--disable-setuid-sandbox",
                    "--ignore-ICertificatePolicy-errors",
                }
            }).ConfigureAwait(false);
            //for (int i = 0; i < 5; i++)
            //{
            //    await _browser.PagesAsync().ConfigureAwait(false);

            //}
            return (await _browser.PagesAsync().ConfigureAwait(false)).First();
        }

        public static async Task<string> GetPageContent(Page page)
        {
            return await page.GetContentAsync().ConfigureAwait(false);
        }
        public static string GetCurrentPageUrl(Page page)
        {
            return page.Url;
        }

        public static async Task<string> GetHtmlContentFromUrl(string amazonUrl, Page page, bool isAmazone = false)
        {
            var amazonContent = "";
            try
            {
                if (!isAmazone)
                {
                    return Helper.GetContentFromUrl(amazonUrl).DocumentNode.OuterHtml;
                }
                else
                {

                    await page.GoToAsync(amazonUrl, new NavigationOptions()
                    {
                        Timeout = 0,
                        WaitUntil = new WaitUntilNavigation[]
                        {
                            WaitUntilNavigation.DOMContentLoaded
                        }
                    }).ConfigureAwait(false);

                }
                amazonContent = await GetPageContent(page);
            }
            catch (Exception e)
            {
                return await GetHtmlContentFromUrl(amazonUrl, page, isAmazone);
            }

            return amazonContent;
        }

        //public static string GetApplyLink(string url,Page page)
        //{
        //    var returnVal = "";

        //    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
        //    doc.LoadHtml(GetHtmlContentFromUrl(url, page).Result);
        //    var elemt = doc.DocumentNode.QuerySelector("#applyButtonLinkContainer a");
        //    if (elemt == null)
        //    {
        //        returnVal = "";
        //    }
        //    returnVal = elemt?.GetAttributeValue("href", null) ?? "";

        //    if (string.IsNullOrEmpty(returnVal))
        //    {
        //        returnVal = "Application Form";
        //    }

        //    return returnVal;
        //}

        public static string GetApplyLink(string url, int tryiNdex)
        {
            HtmlAgilityPack.HtmlDocument doc = Helper.GetContentFromUrl(url);
            var elemt = doc.DocumentNode.QuerySelector("#applyButtonLinkContainer a");

            var returnVal = elemt?.GetAttributeValue("href", null) ?? "";
            if (string.IsNullOrEmpty(returnVal) && tryiNdex < 3)
            {
                returnVal = GetApplyLink(url, ++tryiNdex);
            }
            return returnVal;
        }
    }
}
