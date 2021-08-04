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


                    List<Task> tss = new List<Task>();

                    var sick = chkbxLoadIndeedInBrowser.Checked;
                    var ts = Task.Run(() =>
                    {
                        workd(sick).Wait();

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
                    dataSet = reader.AsDataSet();
                }
            }

            return dataSet;
        }
        private async Task workd(bool ischked)
        {
            var page = BrowserAutoBot.setupBrowser().Result;
            foreach (var currentBatch in jobs.ToList().Skip(Convert.ToInt32(numericUpDown1.Value)).Take(Convert.ToInt32(numericUpDown2.Value) - Convert.ToInt32(numericUpDown1.Value)))
            {
                var batch = currentBatch;
                Invoke((Action)(() => { label2.Text = $@"Processing {jobs.IndexOf(batch) + 1} out of {jobs.Count - 1}"; }));
                await Search(currentBatch, page, ischked).ConfigureAwait(false);
                currentBatch.isProcessed = workStatus.completed;
            }

            Invoke((Action)(() =>
            {
                //MessageBox.Show("Jobs done");
                label2.Text = @"Job Done!! Export manually";
            }));
        }

        private static xlData GetCurrentBatchJob()
        {
            xlData currentBatch;

            currentBatch =
                jobs.FirstOrDefault(x => x.isProcessed == workStatus.pending && string.IsNullOrEmpty(x.JobDetailUrl));

            if (currentBatch != null) currentBatch.isProcessed = workStatus.started;

            return currentBatch;
        }

        void ProcessRows()
        {


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
                    continue;
                jobs.Add(xlDataObj);
            }
        }
        private async Task Search(xlData xlDataObj, Page _page, bool useBrowserasBOt)
        {
            try
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
                doc.LoadHtml(await BrowserAutoBot.GetHtmlContentFromUrl(jobUrl, _page, useBrowserasBOt).ConfigureAwait(false));//GetContentFromUrl(jobUrl);
                var jobList = doc.QuerySelectorAll(".result");

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
                        jobTitle = item.QuerySelector(".jobtitle").InnerText;
                    }
                    catch (Exception)
                    {
                        // ignored
                    }

                    var company = item.QuerySelector(".companyOverviewLink")?.InnerText.Replace("\n", "");
                    if (company.ToLower().Contains("amazon"))
                    {
                        var jobDetailUrl = BrowserAutoBot.GetApplyLink($"{IndeedBaseUrl}/viewjob?jk=" + id, 1, _page, true).HandleEmptyUrl();
                        var jobid = await GetAmazonId(jobDetailUrl, _page).ConfigureAwait(false);
                        if (!string.IsNullOrEmpty(jobid))
                        {
                            foreach (var jb in jobs.Where(jb => jb.xlAmazonId == jobid))
                            {
                                jb.JobDetailUrl = BrowserAutoBot.GetCurrentPageUrl(_page);
                                jb.CompanyName = company;
                            }
                            break;
                            if (jobid == xlDataObj.xlAmazonId)
                            {
                                break;
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
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
            table.Columns.Add("Company Name", typeof(string));
            table.Columns.Add("JobID", typeof(string));
            table.Columns.Add("Job URL", typeof(string));

            foreach (var item in jobs)
            {
                table.Rows.Add(item.xlDate,
                    item.xlSite,
                    item.xlKeyword,
                    item.xlJobLocation,
                    string.IsNullOrEmpty(item.CompanyName) ? "No Job Found" : item.CompanyName,
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
        public string CompanyName { get; set; }
    }
    public enum workStatus
    {
        started,
        pending,
        completed
    }
}
