using System;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using HtmlAgilityPack;
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

        public static async Task<string> GetHtmlContentFromUrl(string amazonUrl, Page page, bool loadUsingBrowserBot = false)
        {
            var amazonContent = "";
            try
            {
                if (!loadUsingBrowserBot)
                {
                    return GetContentFromUrl(amazonUrl).DocumentNode.OuterHtml;
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
                return await GetHtmlContentFromUrl(amazonUrl, page, loadUsingBrowserBot);
            }

            return amazonContent;
        }

        public static string GetApplyLink(string url, int tryiNdex, Page page, bool useBrowserBot = false)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(GetHtmlContentFromUrl(url, page, useBrowserBot).Result);
            var elemt = doc.DocumentNode.QuerySelector("#applyButtonLinkContainer a");

            var returnVal = elemt?.GetAttributeValue("href", null) ?? "";
            if (string.IsNullOrEmpty(returnVal) && tryiNdex < 3)
            {
                returnVal = GetApplyLink(url, ++tryiNdex, page, useBrowserBot);
            }
            return returnVal;
        }

        private static HtmlDocument GetContentFromUrl(string url, int ttry = 1)
        {
            var str = "";

            try
            {

                //using (WebClient wc = new WebClient())
                //{
                //    wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36");
                //    str = wc.DownloadString(url);
                //}
                var web = new HtmlWeb
                {
                    UserAgent =
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
                };
                var doc = web.Load(url);
                return doc;

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Thread.Sleep(5000);
                if (ttry < 3)
                {
                    return GetContentFromUrl(url, ++ttry);
                }
            }
            HtmlDocument d = new HtmlDocument();
            d.LoadHtml(str);
            return d;
        }

    }
}
