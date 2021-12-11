using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PuppeteerSharp;
using Shared;

namespace Indeed_All_Job_Page_Crawler
{
    public partial class Form1 : Form
    {
        private string _indeedBaseUrl = "https://www.indeed.com";
        static string sFileName;
        private static Browser _browser;
        static List<xlData> jobs = new List<xlData>();
        static List<xlData> OuputJobs = new List<xlData>();
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
            public string Position { get; set; }
            public string Company { get; set; }
            public string JobTitle { get; set; }
            public string Location2 { get; set; }
            public string Wage { get; set; }
            public string Age { get; set; }
            public string JobId { get; set; }
            public string JobDesctionLastLine { get; internal set; }
        }
        void PrepareRows(DataSet dataSet)
        {
            var datatable = dataSet.Tables[0];
            for (var iRow = 0; iRow < datatable.Rows.Count; iRow++) // START FROM THE SECOND ROW.
            {
                xlData xlDataObj = new xlData();

                if (datatable.Rows[iRow][1] == null)
                {
                    return;
                }

                xlDataObj.xlDate = datatable.Rows[iRow][0].ToString();
                xlDataObj.xlSite = datatable.Rows[iRow][1].ToString();
                xlDataObj.xlKeyword = datatable.Rows[iRow][2].ToString().RegexReplaceCaseInsenstive("empty", "");
                xlDataObj.xlJobLocation = datatable.Rows[iRow][3].ToString();
                xlDataObj.xlAmazonId = datatable.Rows[iRow][4].ToString();
                xlDataObj.JobDetailUrl = datatable.Rows[iRow][5].ToString();

                xlDataObj.xlJobLocation = xlDataObj.xlJobLocation.RegexReplaceCaseInsenstive("empty", "");

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
                    Task.Run(Work);
                }
            }
        }

        void Work()
        {
            DataSet dataSet = readExcel(sFileName);
            PrepareRows(dataSet);
            var page = setupBrowser().Result;
            var detailPage = setupBrowser().Result;
            foreach (var xlDataObj in jobs)
            {
                string jobUrl = $"{_indeedBaseUrl}/jobs?radius=5";
                if (!string.IsNullOrEmpty(xlDataObj.xlKeyword) && !string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "&q=" + xlDataObj.xlKeyword + "&l=" + xlDataObj.xlJobLocation;
                }
                else if (!string.IsNullOrEmpty(xlDataObj.xlKeyword))
                {
                    jobUrl = jobUrl + "&q=" + xlDataObj.xlKeyword;
                }
                else if (!string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "&l=" + xlDataObj.xlJobLocation;
                }

                page.GoToAsync(jobUrl, new NavigationOptions() { Timeout = 0, WaitUntil = new WaitUntilNavigation[] { } });

                Task.Delay(10000).Wait();
                Invoke((Action)(() => { label1.Text = $@"{jobs.IndexOf(xlDataObj)} processing"; }));
                var htmlCnt = page.GetContentAsync().Result;
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(htmlCnt);

                var jobList = doc.QuerySelectorAll(".jobsearch-SerpJobCard.unifiedRow.row");
                var isMosiac = false;
                if (jobList.Count == 0)
                {
                    isMosiac = true;
                    jobList = doc.QuerySelectorAll("#mosaic-provider-jobcards>a.result");
                }
                List<string> jobIds = new List<string>();
                foreach (var item in jobList)
                {
                    var id = item.Id.Split('_')[item.Id.Split('_').Length - 1];
                    jobIds.Add(id);
                }

                var allJobsDetails = jobDetails(jobIds);

                for (var index = 0; index < jobList.Count; index++)
                {
                    var item = jobList[index];
                    var id = item.Id.Split('_')[item.Id.Split('_').Length - 1];
                    var jobTitle = "";
                    var company = "";
                    var jobId = "";
                    var jobAge = "";
                    var jobWage = "";
                    var jobDetailUrl = "";
                    var jobLocation = "";
                    try
                    {
                        if (isMosiac)
                        {
                            jobTitle = item.QuerySelectorAll(".jobTitle>span").First(x => x.Name == "span").Attributes.FirstOrDefault(x => x.Name.ToLower() == "title")?.Value.HtmlDecode();
                        }
                        else
                        {
                            jobTitle = item.QuerySelector("h2 a").Attributes.FirstOrDefault(x => x.Name.ToLower() == "title")?.Value.HtmlDecode();
                        }
                    }
                    catch (Exception)
                    {
                        // ignored
                    }

                    if (isMosiac)
                    {
                        Invoke((Action)(() => { label2.Text = $@"getting company content : index {index}"; }));
                        company = item.QuerySelector(".companyName")?.InnerText.Replace("\n", "").HtmlDecode();
                        Invoke((Action)(() => { label2.Text = $@"getting detail url : index {index}"; }));
                        jobDetailUrl = item.GetAttributeValue<string>("href", "");
                        if (!string.IsNullOrEmpty(jobDetailUrl))
                        {
                            jobDetailUrl = HttpUtility.HtmlDecode($"{_indeedBaseUrl}{jobDetailUrl}");
                        }
                        else
                        {
                            jobDetailUrl = "Not Found";
                        }
                        // jobId = GetAmazonId(jobDetailUrl, detailPage).Result;
                        Invoke((Action)(() => { label2.Text = $@"getting job wage : index {index}"; }));
                        jobWage = (item.QuerySelector(".metadata.salary-snippet-container")?.InnerText.Trim() ?? "").HandleJobWageFromIndeed();
                        Invoke((Action)(() => { label2.Text = $@"getting job age : index {index}"; }));
                        jobAge = item.QuerySelector(".date").InnerText.HandleStringDateFromIndeed();
                        Invoke((Action)(() => { label2.Text = $@"getting location : index {index}"; }));
                        jobLocation = item.QuerySelector(".companyLocation").ChildNodes[0].InnerText.HandleStringJobLocationFromIndeed().HtmlDecode();

                    }
                    else
                    {
                        Invoke((Action)(() => { label2.Text = $@"getting company content : index {index}"; }));
                        company = item.QuerySelector(".company")?.InnerText.Replace("\n", "").HtmlDecode();
                        Invoke((Action)(() => { label2.Text = $@"getting detail url : index {index}"; }));
                        jobDetailUrl = BrowserAutoBot.GetApplyLink($"{_indeedBaseUrl}/viewjob?jk=" + id, 3, detailPage, false)
                            .HandleEmptyUrl();
                        //jobId = GetAmazonId(jobDetailUrl, detailPage).Result;
                        Invoke((Action)(() => { label2.Text = $@"getting job wage : index {index}"; }));
                        jobWage = (item.QuerySelector(".salaryText")?.InnerText.Trim() ?? "").HandleJobWageFromIndeed();
                        Invoke((Action)(() => { label2.Text = $@"getting job age : index {index}"; }));
                        jobAge = item.QuerySelector(".date").InnerText.HandleStringDateFromIndeed();
                        Invoke((Action)(() => { label2.Text = $@"getting location : index {index}"; }));
                        jobLocation = item.QuerySelector(".recJobLoc").Attributes["data-rc-loc"].Value.HandleStringJobLocationFromIndeed().HtmlDecode();
                    }

                    //detailPage.GoToAsync(jobDetailUrl, new NavigationOptions() { Timeout = 0, WaitUntil = new WaitUntilNavigation[] { } });
                    //Task.Delay(10000).Wait();
                    //var htmlDescCnt = detailPage.GetContentAsync().Result;

                    var compListFilter = new List<string>()
                    {
                        "amazon", "amazon hvh", "amazon workforce staffing"
                    };
                    var des = "";
                    if (compListFilter.Any(x=>x== company.ToLower()))
                    {

                        HtmlAgilityPack.HtmlDocument docDesc = new HtmlAgilityPack.HtmlDocument();
                        docDesc.LoadHtml(allJobsDetails[jobList[index].Id.Split('_')[item.Id.Split('_').Length - 1]]);
                        des = docDesc.DocumentNode.InnerText;
                        var spliter = new string[] { ". " };
                        if (company.ToLower()=="amazon")
                        {
                            spliter = new string[] { "/." };
                        }
                        var desList = des.Split(spliter, StringSplitOptions.RemoveEmptyEntries);

                        des = desList[desList.Length - 1].Trim();
                    }
                    
                    OuputJobs.Add(new xlData()
                    {
                        xlDate = xlDataObj.xlDate,
                        xlKeyword = xlDataObj.xlKeyword,
                        xlJobLocation = xlDataObj.xlJobLocation,
                        xlSite = xlDataObj.xlSite,
                        Company = company,
                        JobTitle = jobTitle,
                        Position = (index + 1).ToString(),
                        JobDetailUrl = jobDetailUrl,
                        JobId = jobId,
                        Age = jobAge,
                        Wage = jobWage,
                        Location2 = jobLocation,
                        JobDesctionLastLine = des
                    });
                }
            }

            ExportToExcel(OuputJobs);
            MessageBox.Show("Processed");
        }


        private async Task<string> GetAmazonId(string JobDetailUrl, Page _page)
        {
            try
            {
                if (JobDetailUrl != "Application Form")
                {
                    await BrowserAutoBot.GetHtmlContentFromUrl(JobDetailUrl, _page, true).ConfigureAwait(false);
                    var amazonContent = await BrowserAutoBot.GetAmazonJobDetailIdPageContent(_page);
                    var amazonId = Helper.GetAmazonJobId(amazonContent);
                    //var tried = 0;
                    //while (amazonId == "" && tried < 5)
                    //{
                    //    Thread.Sleep(5000);
                    //    tried++;
                    //    amazonId = Helper.GetAmazonJobId(await BrowserAutoBot.GetPageContent(_page));
                    //}

                    return amazonId;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                //throw;
            }
            return "";
        }

        private DataTable ExportToExcel(List<xlData> jobs)
        {
            DataSet ds = new DataSet("New_DataSet");
            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("Date", typeof(string));
            table.Columns.Add("Site", typeof(string));
            table.Columns.Add("Search term", typeof(string));
            table.Columns.Add("Search Location", typeof(string));
            table.Columns.Add("Position", typeof(string));
            table.Columns.Add("Company", typeof(string));
            table.Columns.Add("Job Title", typeof(string));
            table.Columns.Add("Location2", typeof(string));
            table.Columns.Add(" Wage ", typeof(string));
            table.Columns.Add("Age", typeof(string));
            table.Columns.Add("Position URL", typeof(string));
            table.Columns.Add("JobID", typeof(string));
            table.Columns.Add("Job Desc Last Line", typeof(string));
            foreach (var item in jobs)
            {
                table.Rows.Add(item.xlDate,
                    item.xlSite,
                    item.xlKeyword,
                    item.xlJobLocation,
                    item.Position,
                    item.Company ?? "",
                item.JobTitle ?? "",
                    item.Location2 ?? "",
                    item.Wage ?? "",
                    item.Age ?? "",
                    item.JobDetailUrl ?? "",
                    item.JobId ?? "",
                    item.JobDesctionLastLine ?? ""
                );
            }

            ds.Tables.Add(table);
            var id = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = @"C:/job/" + id;

            ExcelLibrary.DataSetHelper.CreateWorkbook(path + ".xls", ds);
            return table;
        }

        private DataSet readExcel(string sFile)
        {
            DataSet dataSet;

            using (var stream = File.Open(sFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    dataSet = reader.AsDataSet();
                }
            }

            return dataSet;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(OuputJobs);
        }


        private Dictionary<string, string> jobDetails(List<string> jobIds)
        {
            string delimiter = ",";
            var keywords = String.Join(delimiter, jobIds);
            var url = $"{_indeedBaseUrl}/rpc/jobdescs?jks=" + keywords;
            var html = "";

            using (WebClient wc = new WebClient())
            {
                wc.Headers["accept"] = "application/json";
                wc.Headers["UserAgent"] =
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36";
                Console.WriteLine("downloading-> " + url);
                html = wc.DownloadString(url);
                return JsonConvert.DeserializeObject<Dictionary<string, string>>(html);
            }
        }


    }

}
