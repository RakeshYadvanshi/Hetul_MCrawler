using ExcelDataReader;
using Newtonsoft.Json;
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

namespace Job2Careers_All_Job_Crawl
{


    public partial class Form1 : Form
    {

        #region classes


        public class JobTwoCrawler
        {
            public int total { get; set; }
            public int start { get; set; }
            public int count { get; set; }
            public Dictionary<string, JobDetail> jobAds { get; set; }
        }

        public class SalaryDetail
        {
            public Dictionary<string, string> value { get; set; }
            public string unit { get; set; }
            public string label { get; set; }
        }



        public class JobDetail
        {
            public string id { get; set; }
            public string title { get; set; }
            public DateTime datePosted { get; set; }
            public object onClickSnippet { get; set; }
            public string link { get; set; }
            public string companyName { get; set; }
            public string city { get; set; }
            public string state { get; set; }
            public string cityState { get; set; }
            public string price { get; set; }
            public int preview { get; set; }
            public string description { get; set; }
            public string imageUrl { get; set; }
            public Dictionary<string, SalaryDetail> salaryDetails { get; set; }
            public string primaryMajorCategory { get; set; }
            public string primaryMinorCategory { get; set; }
            public string secondaryMajorCategory { get; set; }
            public string secondaryMinorCategory { get; set; }
            public string adType { get; set; }
            public bool recommended { get; set; }
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
        }

        #endregion
        private readonly string _jobs2careersUrl = "https://j2cweb-backend-prod.jobs2careers.com";
        static string sFileName;
        static List<xlData> jobs = new List<xlData>();
        static List<xlData> OuputJobs = new List<xlData>();


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
                Task.Run(Work);
            }
        }


        private void Work()
        {
            DataSet dataSet = readExcel(sFileName);
            PrepareRows(dataSet);

            foreach (var xlDataObj in jobs)
            {
                List<string> jobIds = new List<string>();
                string jobUrl = $"{_jobs2careersUrl}/api/v1/jobAds/result?sort=r&start=0&categoryId=&jobType=1,2,4&exactMatch=0&distance=15";
                var keywrd = xlDataObj.xlKeyword.ToLower().Replace("empty", "");
                if (!string.IsNullOrEmpty(keywrd) && !string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "&q=" + keywrd + "&l=" + xlDataObj.xlJobLocation;
                }
                else if (!string.IsNullOrEmpty(keywrd))
                {
                    jobUrl = jobUrl + "&q=" + keywrd;
                }
                else if (!string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "&l=" + xlDataObj.xlJobLocation;
                }
                var xs = BrowserAutoBot.GetStringContentFromUrl(jobUrl).Result;

                Invoke((Action)(() =>
                {
                    label1.Text = $@"{jobs.IndexOf(xlDataObj)} Processing..";

                }));


                var output = JsonConvert.DeserializeObject<JobTwoCrawler>(xs);

                Thread.Sleep(10000);
                if (output.jobAds.Keys.Count > 0)
                {
                    foreach (var job in output.jobAds.Values)
                    {
                        OuputJobs.Add(new xlData()
                        {
                            xlDate = xlDataObj.xlDate,
                            xlKeyword = xlDataObj.xlKeyword,
                            xlJobLocation = xlDataObj.xlJobLocation,
                            xlSite = xlDataObj.xlSite,
                            Company = job.companyName,
                            JobTitle = job.title,
                            Position = ((output.jobAds.Values.ToList().IndexOf(job)) + 1).ToString(),
                            JobDetailUrl = job.link,
                            JobId = "",
                            Age = Convert.ToInt32((DateTime.Now - job.datePosted).TotalDays).ToString(),
                            Location2 = job.cityState,
                            Wage = getSalValue(job.salaryDetails),
                            xlAmazonId = ""
                        });
                    }
                }
                else
                {
                    OuputJobs.Add(new xlData()
                    {
                        xlDate = xlDataObj.xlDate,
                        xlKeyword = xlDataObj.xlKeyword,
                        xlJobLocation = xlDataObj.xlJobLocation,
                        xlSite = xlDataObj.xlSite,
                        Company = "No Job Found",
                        JobTitle = "No Job Found",
                        Position = "No Job Found",
                        JobDetailUrl = "",
                        JobId = "",
                        Age = "",
                        Location2 = "No Job Found",
                        Wage = "",
                        xlAmazonId = ""
                    });
                }


            }
            ExportToExcel(OuputJobs);
            Invoke((Action)(() =>
            {
                label1.Text = $@"Exported Successfuly";

            }));
        }

        private string getSalValue(Dictionary<string, SalaryDetail> jd)
        {
            var hourlyValue = jd?.Values.FirstOrDefault(x => x.unit == "hourly")?.value?.Values;
            return hourlyValue == null ? "" : "$"+hourlyValue.Last();
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
                    item.JobId ?? ""

                );
            }

            ds.Tables.Add(table);
            var id = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = @"C:/job/" + id;

            ExcelLibrary.DataSetHelper.CreateWorkbook(path + ".xls", ds);
            return table;
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
                xlDataObj.xlKeyword = datatable.Rows[iRow][2].ToString();
                xlDataObj.xlJobLocation = datatable.Rows[iRow][3].ToString();
                xlDataObj.xlAmazonId = datatable.Rows[iRow][4].ToString();
                xlDataObj.JobDetailUrl = datatable.Rows[iRow][5].ToString();
                if (xlDataObj.xlJobLocation.ToLower().Contains("empty"))
                {
                    xlDataObj.xlJobLocation = xlDataObj.xlJobLocation.ToLower().Replace("empty", "");
                }
                if (string.IsNullOrEmpty(xlDataObj.xlKeyword) && string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                    continue;
                jobs.Add(xlDataObj);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(OuputJobs);
        }
    }
}
