using ExcelDataReader;
using PuppeteerSharp;
using Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ziprecruiter_all_job_crawl
{
    public partial class Form1 : Form
    {
        string _zipRecruiterUrl = "https://www.ziprecruiter.com";
        static string sFileName;
        static List<xlData> jobs = new List<xlData>();
        static List<xlData> ResultJobs = new List<xlData>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = @"Excel File to Edit";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = @"Excel File|*.xlsx;*.xls";

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

            Invoke((Action)(() => label2.Text = @"Job Done!! Export manually"));
        }

        private async Task Search(xlData xlDataObj, Page _page, bool useBrowserAsBot)
        {
            try
            {

                List<string> jobIds = new List<string>();
                string jobUrl = $"{_zipRecruiterUrl}/candidate/search";

                if (!string.IsNullOrEmpty(xlDataObj.xlKeyword) && !string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "?search=" + xlDataObj.xlKeyword + "&location=" + xlDataObj.xlJobLocation;
                }
                else if (!string.IsNullOrEmpty(xlDataObj.xlKeyword))
                {
                    jobUrl = jobUrl + "?search=" + xlDataObj.xlKeyword;
                }
                else if (!string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "?location=" + xlDataObj.xlJobLocation;
                }
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(await BrowserAutoBot.GetHtmlContentFromUrl(jobUrl, _page, true).ConfigureAwait(false));

                var jobList = doc.QuerySelectorAll(".job_result.t_job_result");

                for (var index = 0; index < jobList.Count; index++)
                {
                    var item = jobList[index];

                    var jobTitle = item.QuerySelector(".just_job_title")?.InnerText.Replace("\n", "");

                    var company = item.QuerySelector(".t_org_link.name")?.InnerText.Replace("\n", "");
                    var jobDetailUrl = item.QuerySelector(".job_link.t_job_link").Attributes["href"].Value;

                    ResultJobs.Add(new xlData
                    {
                        CompanyName = xlDataObj.CompanyName,
                        xlJobLocation = xlDataObj.xlJobLocation,
                        xlKeyword = xlDataObj.xlKeyword,
                        xlDate = xlDataObj.xlDate,
                        xlSite = xlDataObj.xlSite,
                        Position = index + 1,
                        Company = company,
                        JobTitle = jobTitle,
                        Location2 = "",
                        Wage = "NA",
                        Age = "NA",
                        JobDetailUrl = jobDetailUrl
                    });

                    //if (company.ToLower().Contains("amazon"))
                    //{
                    //    var jobid = await GetAmazonId(jobDetailUrl, _page).ConfigureAwait(false);
                    //    if (!string.IsNullOrEmpty(jobid))
                    //    {
                    //        foreach (var jb in jobs.Where(jb => jb.xlAmazonId == jobid))
                    //        {
                    //            jb.JobDetailUrl = BrowserAutoBot.GetCurrentPageUrl(_page);
                    //            jb.CompanyName = company;
                    //        }
                    //        break;
                    //    }

                    //}

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
        }

        private async Task<string> GetAmazonId(string jobDetailUrl, Page _page)
        {
            try
            {
                if (jobDetailUrl != "Application Form")
                {
                    var amazonContent = await BrowserAutoBot.GetHtmlContentFromUrl(jobDetailUrl, _page, true).ConfigureAwait(false);
                    var amazonId = Helper.GetAmazonJobId(amazonContent);
                    var tried = 0;
                    while (amazonId == "" && tried < 2)
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
            }
            return "";
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
            table.Columns.Add("Job URL", typeof(string));
            table.Columns.Add("JobID", typeof(string));

            foreach (var item in jobs)
            {
                table.Rows.Add(item.xlDate,
                    item.xlSite,
                    item.xlKeyword,
                    item.xlJobLocation,
                    item.Position,
                    string.IsNullOrEmpty(item.CompanyName) ? "No Job Found" : item.CompanyName,
                    item.JobTitle,
                    item.Location2,
                    item.Wage,
                    item.Age,
                    item.JobDetailUrl ?? "",
                    item.xlAmazonId
                );
            }

            ds.Tables.Add(table);
            var id = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = @"C:/job/" + id;

            ExcelLibrary.DataSetHelper.CreateWorkbook(path + ".xls", ds);
            return table;
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
            public int Position { get; internal set; }
            public string JobTitle { get; internal set; }
            public string Company { get; internal set; }
            public string Location2 { get; internal set; }
            public string Wage { get; internal set; }
            public string Age { get; internal set; }
        }
        public enum workStatus
        {
            started,
            pending,
            completed
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(ResultJobs);
        }


    }
}
