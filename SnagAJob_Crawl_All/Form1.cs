﻿using ExcelDataReader;
using Newtonsoft.Json;
using Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SnagAJob_Crawl_All
{


    public partial class Form1 : Form
    {


        #region SnapAJobClass

        public class SnagajobClass
        {

            public List[] list { get; set; }

        }


        public class List
        {

            public string postingId { get; set; }
            public string companyName { get; set; }
            public string title { get; set; }

            public DateTime createdDate { get; set; }

            public AddressLocation location { get; set; }

        }

        public class AddressLocation
        {

            public string city { get; set; }

            public string stateProvinceCode { get; set; }

        }



        #endregion
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
            public string AmazonLink { get; set; }
        }
        private readonly string _snagAJobUrl = "https://www.snagajob.com";
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



        void Work()
        {
            DataSet dataSet = readExcel(sFileName);
            PrepareRows(dataSet);

            foreach (var xlDataObj in jobs)
            {
                List<string> jobIds = new List<string>();
                string jobUrl = $"{_snagAJobUrl}/api/jobs/v1/p4p?radiusInMiles=5&num=15";
                if (!string.IsNullOrEmpty(xlDataObj.xlKeyword) && !string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "&query=" + xlDataObj.xlKeyword + "&location=" + xlDataObj.xlJobLocation;
                }
                else if (!string.IsNullOrEmpty(xlDataObj.xlKeyword))
                {
                    jobUrl = jobUrl + "&query=" + xlDataObj.xlKeyword;
                }
                else if (!string.IsNullOrEmpty(xlDataObj.xlJobLocation))
                {
                    jobUrl = jobUrl + "&location=" + xlDataObj.xlJobLocation;
                }

                Invoke((Action)(() =>
                {
                    label1.Text = $@"{jobs.IndexOf(xlDataObj)} Processing..";
                }));
                var output = JsonConvert.DeserializeObject<SnagajobClass>(BrowserAutoBot.GetStringContentFromUrl(jobUrl).Result);

                Thread.Sleep(10000);
                if (output.list.Length > 0)
                {
                    foreach (var job in output.list)
                    {
                        var amazonLink = "";
                        if (job.companyName.ToLower() == "amazon")
                        {
                            amazonLink = "https://www.snagajob.com/job-seeker/apply/apply.aspx?postingId=" +
                                         job.postingId;
                        }
                        OuputJobs.Add(new xlData()
                        {
                            xlDate = xlDataObj.xlDate,
                            xlKeyword = xlDataObj.xlKeyword,
                            xlJobLocation = xlDataObj.xlJobLocation,
                            xlSite = xlDataObj.xlSite,
                            Company = job.companyName,
                            JobTitle = job.title,
                            Position = ((output.list.ToList().IndexOf(job)) + 1).ToString(),
                            JobDetailUrl = $"{_snagAJobUrl}/jobs/{job.postingId}",
                            JobId = "",
                            Age = ((int)(DateTime.Now - job.createdDate).TotalDays) + " days",
                            Location2 = job.location?.city + ", " + job.location?.stateProvinceCode,
                            Wage = "",
                            xlAmazonId = "",
                            AmazonLink = amazonLink
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
            MessageBox.Show("Processed");
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
                xlDataObj.xlKeyword = datatable.Rows[iRow][2].ToString();
                if (datatable.Rows[iRow][2].ToString().ToLower().Contains("empty"))
                {
                    xlDataObj.xlKeyword = datatable.Rows[iRow][2].ToString().ToLower().Replace("empty", "");
                }

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
            MessageBox.Show("Processed");
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
            table.Columns.Add("Amazon Link", typeof(string));
            foreach (var item in jobs)
            {
                table.Rows.Add(item.xlDate,
                    item.xlSite,
                    item.xlKeyword,
                    item.xlJobLocation,
                    item.Position,
                    item.Company ?? "",
                    item.JobTitle ?? "",
                    (item.Location2 ?? "").ToCamelCase(),
                    item.Wage ?? "",
                    item.Age ?? "",
                    item.JobDetailUrl ?? "",
                    item.JobId ?? "",
                    item.AmazonLink ?? ""

                );
            }

            ds.Tables.Add(table);
            var id = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = @"C:/job/" + id;

            ExcelLibrary.DataSetHelper.CreateWorkbook(path + ".xls", ds);
            return table;
        }
    }
}