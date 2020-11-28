using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using Newtonsoft.Json;
using PuppeteerSharp;
using Shared;

namespace IndeedCrawler
{
    public partial class Form1 : Form
    {
        private Page _page;
        public Form1()
        {
            InitializeComponent();
            _page = BrowserAutoBot.setupBrowser().Result;
        }


        string IndeedBaseUrl = "https://www.indeed.com";
        static string sFileName;
        static int iRow, iCol = 2;
        static List<xlData> jobs = new List<xlData>();

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

                    Task.Run(() =>
                    {
                        jobs = new List<xlData>();
                        processRows(dataSet);
                        ExportToExcel(jobs);
                        Invoke((Action)(() =>
                        {
                            MessageBox.Show("Jobs done");
                            label2.Text = "";
                        }));

                    });
                }

            }
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
            table.Columns.Add("Wage", typeof(string));
            table.Columns.Add("Age", typeof(string));
            table.Columns.Add("Position URL", typeof(string));
            table.Columns.Add("Amazon JobId", typeof(string));
            table.Columns.Add("Column1", typeof(string));


            foreach (var item in jobs)
            {
                table.Rows.Add(item.xlDate,
                    item.xlSite,
                    item.xlKeyword,
                    item.xlJobLocation,
                    item.xlJobIndex,
                    item.JobCompany ?? "",
                    item.JobTitle,
                    item.JobLocation,
                    item.JobWage ?? "",
                    item.JobAge,
                    item.JobDetailUrl,
                    item.AmazonJobId ?? "",
                    ""
                    );
            }

            ds.Tables.Add(table);
            var id = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = @"C:/job/" + id;

            ExcelLibrary.DataSetHelper.CreateWorkbook(path + ".xls", ds);
            return table;
        }


        void processRows(DataSet dataSet)
        {
            var datatable = dataSet.Tables[0];
            for (iRow = 1; iRow < datatable.Rows.Count; iRow++) // START FROM THE SECOND ROW.
            {
                xlData xlDataObj = new xlData();

                if (datatable.Rows[iRow][1] == null)
                {
                    return;
                }

                Invoke((Action)(() =>
                {
                    label2.Text = $@"Processing {iRow - 1} out of {datatable.Rows.Count - 1}";
                }));
                xlDataObj.xlDate = datatable.Rows[iRow][0].ToString();
                xlDataObj.xlSite = datatable.Rows[iRow][1].ToString();
                xlDataObj.xlKeyword = datatable.Rows[iRow][2].ToString().ToLower().Replace("empty", "");
                xlDataObj.xlJobLocation = datatable.Rows[iRow][3].ToString();
                xlDataObj.xlJobIndex = Convert.ToInt32(datatable.Rows[iRow][4].ToString());
                xlDataObj.xlJobLocation = xlDataObj.xlJobLocation.ToLower().Replace("empty", "");
                if (string.IsNullOrEmpty(xlDataObj.xlKeyword) && string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                    return;
                Search(xlDataObj);
            }
        }


        private void Search(xlData xlDataObj)
        {
            List<string> jobIds = new List<string>();

            string jobUrl = $"{IndeedBaseUrl}/jobs";

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
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(BrowserAutoBot.GetHtmlContentFromUrl(jobUrl, _page).Result);//GetContentFromUrl(jobUrl);
            var jobList = doc.QuerySelectorAll(".jobsearch-SerpJobCard.unifiedRow.row");

            foreach (var item in jobList)
            {
                var id = item.Id.Split('_')[item.Id.Split('_').Length - 1];
                jobIds.Add(id);
            }

            var isLastSaved = false;
            var allJobsDetails = jobDetails(jobIds);
            for (var index = 0; index < jobList.Count; index++)
            {
                var item = jobList[index];
                var id = item.Id.Split('_')[item.Id.Split('_').Length - 1];
                var jobTitle = "";
                try
                {
                    jobTitle = item.QuerySelector("h2 a").Attributes.FirstOrDefault(x => x.Name.ToLower() == "title")?.Value;
                }
                catch (Exception)
                {
                    // ignored
                }

                if (string.IsNullOrEmpty(jobTitle))
                {
                    jobTitle = item.QuerySelector("h2 a").InnerText.Replace("amp;", "").Replace("\n", "");
                }
                var jobLocation = item.QuerySelector(".location").InnerText;


                var jobWage = item.QuerySelector(".salaryText")?.InnerText.Trim() ?? "";
                var company = item.QuerySelector(".company")?.InnerText.Replace("\n", "");

                var jobAge = item.QuerySelector(".date").InnerText;

                xlDataObj.JobCompany = company;
                xlDataObj.JobTitle =WebUtility.HtmlDecode(jobTitle);
                xlDataObj.JobLocation = jobLocation;
                xlDataObj.JobAge = jobAge.HandleStringDateFromIndeed();
                xlDataObj.IsPandologic = allJobsDetails[id].Contains("PandoLogic");

                if (jobWage.Count(x => x == '$') > 1 && jobWage.IndexOf('-') > -1)
                {
                    jobWage = jobWage.Split('-')[1];
                    while (jobWage.IndexOf('+') > -1)
                    {
                        jobWage = jobWage.Replace("+", "");
                    }
                }
                xlDataObj.JobWage = jobWage;
                if (xlDataObj.xlJobIndex == 3)
                {

                    if (!isLastSaved)
                    {


                        if (index == xlDataObj.xlJobIndex - 1)
                        {
                            if (xlDataObj.JobLocation.IndexOf(",", StringComparison.Ordinal) > -1)
                            {
                                var joblocationArray = xlDataObj.JobLocation.Split(',');
                                xlDataObj.JobLocation = joblocationArray[0] + ", " + joblocationArray[1].Trim().Split(' ')[0];
                            }
                            isLastSaved = true;
                            xlDataObj.JobDetailUrl = BrowserAutoBot.GetApplyLink($"{IndeedBaseUrl}/viewjob?jk=" + id,1,_page).HandleEmptyUrl();
                            xlDataObj = updateAmazonId(xlDataObj).Result;
                            jobs.Add(xlDataObj);
                            xlDataObj = JsonConvert.DeserializeObject<xlData>(JsonConvert.SerializeObject(xlDataObj));

                        }

                    }
                    else
                    {
                        if (!jobs.Any(x => x.JobCompany.ToLower().Contains("amazon") && x.xlKeyword == xlDataObj.xlKeyword && x.xlJobLocation == xlDataObj.xlJobLocation && x.IsPandologic))
                            if (xlDataObj.JobCompany.ToLower().Contains("amazon") && xlDataObj.IsPandologic)
                            {
                                if (xlDataObj.JobLocation.IndexOf(",") > -1)
                                {
                                    var joblocationArray = xlDataObj.JobLocation.Split(',');
                                    xlDataObj.JobLocation =
                                        joblocationArray[0] + ", " + joblocationArray[1].Trim().Split(' ')[0];
                                }

                                xlDataObj.xlJobIndex = index + 1;
                                xlDataObj.JobDetailUrl = BrowserAutoBot.GetApplyLink($"{IndeedBaseUrl}/viewjob?jk=" + id,1,_page).HandleEmptyUrl();
                                xlDataObj = updateAmazonId(xlDataObj).Result;
                                jobs.Add(xlDataObj);
                                xlDataObj = JsonConvert.DeserializeObject<xlData>(JsonConvert.SerializeObject(xlDataObj));
                                break;
                            }
                    }


                }
                else if (xlDataObj.xlJobIndex == 2 || xlDataObj.xlJobIndex == 1)
                {
                    if (index == xlDataObj.xlJobIndex - 1)
                    {
                        if (xlDataObj.JobLocation.IndexOf(",", StringComparison.Ordinal) > -1)
                        {
                            var joblocationArray = xlDataObj.JobLocation.Split(',');
                            xlDataObj.JobLocation = joblocationArray[0] + ", " + joblocationArray[1].Trim().Split(' ')[0];
                        }
                        xlDataObj.JobDetailUrl = BrowserAutoBot.GetApplyLink($"{IndeedBaseUrl}/viewjob?jk=" + id,1,_page).HandleEmptyUrl();
                        xlDataObj = updateAmazonId(xlDataObj).Result;
                        jobs.Add(xlDataObj);
                        xlDataObj = JsonConvert.DeserializeObject<xlData>(JsonConvert.SerializeObject(xlDataObj));
                        break;
                    }
                }



            }
        }

        private async Task<xlData> updateAmazonId(xlData xlDataObj)
        {
            try
            {
                if (xlDataObj.JobCompany.ToLower().Contains("amazon") && xlDataObj.JobDetailUrl != "Application Form")
                {
                    var amazonContent = await BrowserAutoBot.GetHtmlContentFromUrl(xlDataObj.JobDetailUrl, _page, true).ConfigureAwait(false);
                    var amazonId = Helper.GetAmazonJobId(amazonContent);
                    var tried = 0;
                    while (amazonId == "" && tried < 5)
                    {
                        Thread.Sleep(5000);
                        tried++;
                        amazonId = Helper.GetAmazonJobId(await BrowserAutoBot.GetPageContent(_page));
                    }
                    xlDataObj.JobDetailUrl = BrowserAutoBot.GetCurrentPageUrl(_page);
                    xlDataObj.AmazonJobId = amazonId;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                //throw;
            }
            return xlDataObj;
        }


        private Dictionary<string, string> jobDetails(List<string> jobIds)
        {
            string delimiter = ",";
            var keywords = String.Join(delimiter, jobIds);
            var url = $"{IndeedBaseUrl}/rpc/jobdescs?jks=" + keywords;
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



        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(jobs);
        }

        public class xlData
        {
            public string xlDate { get; set; }
            public int xlJobIndex { get; set; }
            public string xlJobLocation { get; set; }
            public string xlSite { get; set; }
            public string xlKeyword { get; set; }
            public string JobTitle { get; set; }
            public string JobLocation { get; set; }
            public string JobCompany { get; set; }
            public string JobAge { get; set; }
            public string JobDetailUrl { get; set; }
            public string JobWage { get; set; }
            public bool IsPandologic { get; set; }
            public string AmazonJobId { get; set; }
        }
    }
}
