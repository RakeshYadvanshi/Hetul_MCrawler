using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using Newtonsoft.Json;
using PuppeteerSharp;
using Shared;

namespace AmazonJobComarisionCheck
{

    public partial class Form1 : Form
    {

        string IndeedBaseUrl = "https://www.indeed.com";
        static string sFileName;
        static List<xlData> jobs = new List<xlData>();

        public Form1()
        {
            InitializeComponent();

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

                    jobs = new List<xlData>();
                    PrepareRows(dataSet);
                    Task.Run(() =>
                    {
                        ProcessRows();
                        ExportToExcel(jobs);
                        MessageBox.Show("Jobs done");
                        //label2.Text = "";


                    });

                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(jobs);
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
        private void workd()
        {
            var currentBatch = jobs.Where(x => x.isProcessed == workStatus.pending && string.IsNullOrEmpty(x.JobDetailUrl)).FirstOrDefault();
            var page = BrowserAutoBot.setupBrowser().Result;
            while (currentBatch != null)
            {
                Invoke((Action)(() => { label2.Text = $@"Processing {jobs.IndexOf(currentBatch) + 1} out of {jobs.Count - 1}"; }));
                currentBatch.isProcessed = workStatus.started;
                Search(currentBatch, 1, page).Wait();
                currentBatch.isProcessed = workStatus.completed;
                currentBatch = jobs.Where(x => x.isProcessed == workStatus.pending && string.IsNullOrEmpty(x.JobDetailUrl)).FirstOrDefault();
            }
        }
        void ProcessRows()
        {
            List<Task> tss = new List<Task>();
            var instCount = 5;
            //var pageList = new List<Page>();
            for (int i = 0; i < instCount; i++)
            {

                var ts = Task.Run(async () =>
                          {
                              workd();
                          });
                tss.Add(ts);

                //pageList.Add(page);
            }

            Task.WaitAll(tss.ToArray());

            //while (jobs.Any(x => !x.isProcessed))
            //{
            //    var currentBatch = jobs.Where(x => !x.isProcessed).Take(instCount).ToList();
            //    foreach (var xlData in currentBatch)
            //    {
            //        Invoke((Action)(() => { label2.Text = $@"Processing {jobs.IndexOf(xlData) + 1} out of {jobs.Count - 1}"; }));

            //        if (string.IsNullOrEmpty(xlData.JobDetailUrl))
            //        {
            //            var ts = Task.Run(async () =>
            //              {
            //                  await Search(xlData, 1, pageList[currentBatch.IndexOf(xlData)]);
            //                  xlData.isProcessed = true;
            //                  //await page.CloseAsync().ConfigureAwait(false);
            //              });
            //            tss.Add(ts);

            //        }

            //    }

            //    Task.WhenAll(tss).Wait();
            //    tss.Clear();
            //}
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
                    return;
                jobs.Add(xlDataObj);
            }
        }
        private async Task Search(xlData xlDataObj, int indx, Page _page)
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
            doc.LoadHtml(await BrowserAutoBot.GetHtmlContentFromUrl(jobUrl, _page).ConfigureAwait(false));//GetContentFromUrl(jobUrl);
            var jobList = doc.QuerySelectorAll(".jobsearch-SerpJobCard.unifiedRow.row");

            foreach (var item in jobList)
            {
                var id = item.Id.Split('_')[item.Id.Split('_').Length - 1];
                jobIds.Add(id);
            }

            await jobDetails(jobIds).ConfigureAwait(false);
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

                var company = item.QuerySelector(".company")?.InnerText.Replace("\n", "");
                if (company.ToLower().Contains("amazon"))
                {
                    var jobDetailUrl = BrowserAutoBot.GetApplyLink($"{IndeedBaseUrl}/viewjob?jk=" + id, 1).HandleEmptyUrl();
                    var jobid = await GetAmazonId(jobDetailUrl, _page).ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(jobid))
                    {
                        foreach (var jb in jobs.Where(jb => jb.xlAmazonId == jobid))
                        {
                            jb.JobDetailUrl = BrowserAutoBot.GetCurrentPageUrl(_page);
                        }

                        if (jobid == xlDataObj.xlAmazonId)
                        {
                            break;
                        }
                    }

                }

            }

            //if (string.IsNullOrEmpty(xlDataObj.JobDetailUrl) && indx < 3)
            //{
            //    indx++;
            //    Search(xlDataObj, indx);
            //}
            //else
            //{
            //jobs.Add(xlDataObj);
            // }
        }
        private async Task<string> GetAmazonId(string JobDetailUrl, Page _page)
        {
            try
            {
                if (JobDetailUrl != "Application Form")
                {
                    var amazonContent = await BrowserAutoBot.GetHtmlContentFromUrl(JobDetailUrl, _page, true).ConfigureAwait(false);
                    var amazonId = Helper.GetAmazonJobId(amazonContent);
                    var tried = 0;
                    while (amazonId == "" && tried < 5)
                    {
                        Thread.Sleep(5000);
                        tried++;
                        amazonId = Helper.GetAmazonJobId(await BrowserAutoBot.GetPageContent(_page));
                    }

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

        private async Task<Dictionary<string, string>> jobDetails(List<string> jobIds)
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
                html = await wc.DownloadStringTaskAsync(url).ConfigureAwait(false);
                return JsonConvert.DeserializeObject<Dictionary<string, string>>(html);
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
            table.Columns.Add("JobID", typeof(string));
            table.Columns.Add("Job URL", typeof(string));

            foreach (var item in jobs)
            {
                table.Rows.Add(item.xlDate,
                    item.xlSite,
                    item.xlKeyword,
                    item.xlJobLocation,
                    item.xlAmazonId,
                    item.JobDetailUrl ?? ""
                );
            }

            ds.Tables.Add(table);
            var id = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = @"C:/job/" + id;

            ExcelLibrary.DataSetHelper.CreateWorkbook(path + ".xls", ds);
            return table;
        }

    }
    public class xlData
    {
        public string xlDate { get; set; }
        public string xlAmazonId { get; set; }
        public string xlJobLocation { get; set; }
        public string xlSite { get; set; }
        public string xlKeyword { get; set; }
        public string JobDetailUrl { get; set; }
        public workStatus isProcessed { get; set; } = workStatus.pending;
    }
    public enum workStatus
    {
        started,
        pending,
        completed
    }
}
