using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using PuppeteerSharp;

namespace Indeed_All_Job_Page_Crawler
{
    public partial class Form1 : Form
    {
        private string _indeedBaseUrl = "https://www.indeed.com";
        static string sFileName;
        private static Browser _browser;
        static List<xlData> jobs = new List<xlData>();

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

            return (await _browser.PagesAsync().ConfigureAwait(false)).First();
        }
        public Form1()
        {
            InitializeComponent();
        }

        public class xlData
        {
            public string xlDate { get; set; }
            public string xlAmazonId { get; set; }
            public string xlJobLocation { get; set; }
            public string xlSite { get; set; }
            public string xlKeyword { get; set; }
            public string JobDetailUrl { get; set; }
        }
        void PrepareRows(DataSet dataSet)
        {
            var datatable = dataSet.Tables[0];
            for (var iRow = 1; iRow < datatable.Rows.Count; iRow++) // START FROM THE SECOND ROW.
            {
                xlData xlDataObj = new xlData();

                if (datatable.Rows[iRow][1] == null)
                {
                    return;
                }

                xlDataObj.xlDate = datatable.Rows[iRow][0].ToString();
                xlDataObj.xlSite = datatable.Rows[iRow][1].ToString();
                xlDataObj.xlKeyword = datatable.Rows[iRow][2].ToString().ToLower().Replace("empty", "");
                xlDataObj.xlJobLocation = datatable.Rows[iRow][3].ToString();
                xlDataObj.xlAmazonId = datatable.Rows[iRow][4].ToString();
                xlDataObj.JobDetailUrl = datatable.Rows[iRow][5].ToString();
                xlDataObj.xlJobLocation = xlDataObj.xlJobLocation.ToLower().Replace("empty", "");

                if (string.IsNullOrEmpty(xlDataObj.xlKeyword) && string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                    continue;
                jobs.Add(xlDataObj);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Excel File to Edit";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog1.FileName;

                if (sFileName.Trim() != "")
                {
                    DataSet dataSet = readExcel(sFileName);
                    PrepareRows(dataSet);
                    foreach (var xlDataObj in jobs)
                    {
                        List<string> jobIds = new List<string>();

                        string jobUrl = $"{_indeedBaseUrl}/jobs";
                        if (!string.IsNullOrEmpty(xlDataObj.xlKeyword) && !string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                        {
                            jobUrl = jobUrl + "?q=" + xlDataObj.xlKeyword + "&l=" + xlDataObj.xlJobLocation;
                        }
                        else if (!string.IsNullOrEmpty(xlDataObj.xlKeyword))
                        {
                            jobUrl = jobUrl + "?q=" + xlDataObj.xlKeyword;
                        }
                        else if (!string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                        {
                            jobUrl = jobUrl + "?l=" + xlDataObj.xlJobLocation;
                        }

                        var page = setupBrowser().Result;

                        page.GoToAsync(jobUrl, new NavigationOptions()
                        {
                            Timeout = 0,
                            WaitUntil = new WaitUntilNavigation[]
                            {

                            }
                        });

                        Thread.Sleep(3000);

                        var htmlCnt = page.GetContentAsync().Result;
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(htmlCnt);

                        var jobList = doc.QuerySelectorAll(".jobsearch-SerpJobCard.unifiedRow.row");

                        for (var index = 0; index < jobList.Count; index++)
                        {
                            var item = jobList[index];
                            var id = item.Id.Split('_')[item.Id.Split('_').Length - 1];
                            var jobTitle = "";
                            try
                            {
                                jobTitle = item.QuerySelector("h2 a").Attributes
                                    .FirstOrDefault(x => x.Name.ToLower() == "title")?.Value;
                            }
                            catch (Exception)
                            {
                                // ignored
                            }

                            var company = item.QuerySelector(".company")?.InnerText.Replace("\n", "");
                        }
                    }

                }
            }
        }



        private DataSet readExcel(string sFile)
        {
            DataSet dataSet;

            using (var stream = File.Open(sFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    dataSet = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                }
            }

            return dataSet;
        }
    }

}
