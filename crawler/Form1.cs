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
using HtmlDocument = HtmlAgilityPack.HtmlDocument;


namespace crawler
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string baseUrl = "https://www.indeed.com/";//"https://www.indeed.com/";//
        static string sFileName;
        static int iRow, iCol = 2;
        static List<xlData> jobs = new List<xlData>();


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

        void processRows(DataSet dataSet)
        {
            var datatable = dataSet.Tables[0];
            for (iRow = 2; iRow < datatable.Rows.Count; iRow++) // START FROM THE SECOND ROW.
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
                xlDataObj.xlJobLocation = xlDataObj.xlJobLocation.ToLower().Replace("empty", "");

                if (string.IsNullOrEmpty(xlDataObj.xlKeyword) && string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                    return;
                Search(xlDataObj);
            }
        }

        private class xlData
        {
            public string xlDate { get; set; }
            public string xlSite { get; set; }
            public string xlJobLocation { get; set; }
            public string xlKeyword { get; set; }
            public int JobPosition { get; internal set; }
            public string JobCompany { get; set; }
            public string JobTitle { get; set; }
            public string JobLocation { get; internal set; }
            public string JobWage { get; internal set; }

            public string JobAge { get; set; }
            public string DetailUrl { get; set; }
            public string SearchUrl { get; set; }

            public string JobId { get; set; }
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

        private HtmlAgilityPack.HtmlDocument GetContentFromUrl(string url)
        {
            HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb
            {
                UserAgent =
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
            };
            HtmlAgilityPack.HtmlDocument doc = web.Load(url);
            return doc;
        }

        private void Search(xlData xlDataObj)
        {
            try
            {
                List<string> jobIds = new List<string>();

                string jobUrl = baseUrl + "jobs";

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
                xlDataObj.SearchUrl = jobUrl;
                HtmlDocument doc = GetContentFromUrl(jobUrl);
                var jobList = doc.QuerySelectorAll(".jobsearch-SerpJobCard.unifiedRow.row");

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
                    var company = item.QuerySelector(".company")?.InnerText.Replace("\n", "");
                    var jobAge = item.QuerySelector(".date").InnerText.Replace("days ago", "").Replace("day ago", "").Replace("Just Posted", "0").Replace("Today", "0");

                    var jobLocation = item.QuerySelector(".location").InnerText;

                    var jobWage = item.QuerySelector(".salaryText")?.InnerText.Trim() ?? "";
                    if (jobWage.Count(x => x == '$') > 1 && jobWage.IndexOf('-') > -1)
                    {
                        jobWage = jobWage.Split('-')[1];
                        while (jobWage.IndexOf('+') > -1)
                        {
                            jobWage = jobWage.Replace("+", "");
                        }
                    }

                    xlDataObj.JobCompany = company;
                    bool containsPandoLogicWord = allJobsDetails[id].Contains("PandoLogic");

                    if (containsPandoLogicWord && xlDataObj.JobCompany?.ToLower().Trim() == "amazon")
                    {
                        xlDataObj.JobWage = jobWage;
                        xlDataObj.JobTitle = jobTitle;
                        xlDataObj.JobCompany = company;
                        xlDataObj.JobAge = jobAge;
                        xlDataObj.JobLocation = jobLocation;
                        if (xlDataObj.JobLocation.IndexOf(",", StringComparison.Ordinal) > -1)
                        {
                            var joblocationArray = xlDataObj.JobLocation.Split(',');
                            xlDataObj.JobLocation = joblocationArray[0] + ", " + joblocationArray[1].Trim().Split(' ')[0];
                        }
                        xlDataObj.JobPosition = index + 1;
                        xlDataObj.DetailUrl  = GetApplyLink(baseUrl + "viewjob?jk=" + id);
                        if (string.IsNullOrEmpty(xlDataObj.DetailUrl))
                        {
                            xlDataObj.DetailUrl = "Application Form";
                        }
                        break;
                    }
                }

                jobs.Add(xlDataObj);
            }
            catch (Exception e)
            {
                Invoke((Action)(() =>
               {
                   MessageBox.Show(e.ToString());
               }));
                throw e;
            }

        }

        private DataTable ExportToExcel(List<xlData> jobs)
        {
            DataSet ds = new DataSet("New_DataSet");
            DataTable table = new DataTable();
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
            table.Columns.Add("JobID", typeof(string));
            foreach (var item in jobs)
            {
                table.Rows.Add(item.xlDate.ToString(),
                    item.xlSite ?? "",
                    item.xlKeyword ?? "",
                    item.xlJobLocation ?? "",
                    item.JobPosition > 0 ? item.JobPosition.ToString() : "No Job found",
                    item.JobPosition > 0 ? item.JobCompany : "No Job found",
                    item.JobPosition > 0 ? item.JobTitle : "No Job found",
                    item.JobPosition > 0 ? item.JobLocation : "No Job found",
                    item.JobPosition > 0 ? item.JobWage : "No Job found",
                    item.JobPosition > 0 ? item.JobAge : "No Job found",
                    item.JobPosition > 0 ? (item.DetailUrl ?? "") : "No Job found" ,
                    item.JobPosition > 0 ? (item.JobId ?? "") : "No Job found"
                    );
            }

            ds.Tables.Add(table);
            var id = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = @"C:/job/" + id;

            ExcelLibrary.DataSetHelper.CreateWorkbook(path + ".xls", ds);
            return table;
        }


        private Dictionary<string, string> jobDetails(List<string> jobIds)
        {
            string delimiter = ",";
            var keywords = String.Join(delimiter, jobIds);
            var url = baseUrl + "rpc/jobdescs?jks=" + keywords;
            var html = "";

            using (WebClient wc = new WebClient())
            {
                wc.Headers["accept"] = "application/json";
                wc.Headers["UserAgent"] =
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36";
                Console.WriteLine(@"downloading-> " + url);
                html = wc.DownloadString(url);
                return JsonConvert.DeserializeObject<Dictionary<string, string>>(html);
            }
        }

        private string GetApplyLink(string url)
        {
            var returnVal = "";

            HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb
            {
                UserAgent =
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
            };
            HtmlAgilityPack.HtmlDocument doc = web.Load(url);
            var elemt = doc.DocumentNode.QuerySelector("#applyButtonLinkContainer a");
           
            returnVal = elemt?.GetAttributeValue("href", null) ?? "";
            return returnVal;
        }
    }


}
