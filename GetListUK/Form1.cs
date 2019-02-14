using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint;
using System.IO;

namespace GetListUK
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> listColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(txtInput.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    listColl.Add(sr.ReadLine().Trim());
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            StreamWriter ExcelwriterScoringMatrixNew = null;
            ExcelwriterScoringMatrixNew = System.IO.File.CreateText(txtReport.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-mm-yyyy-hh-mm-ss") + ".csv");
            ExcelwriterScoringMatrixNew.WriteLine("SiteURL" +","+ "PageName" +","+ "PageUrl" +","+ "PageId" +","+ "ListType" +","+ "Type" +","+ "webpart" +","+ "url");
            ExcelwriterScoringMatrixNew.Flush();
        }

        private void btninputCSV_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog inputfile = new OpenFileDialog();

            if(inputfile.ShowDialog()==DialogResult.OK)
            {
                txtInput.Text = inputfile.FileName;
            }
            #region Commented
            //inputfile.Multiselect = true;
            //inputfile.ShowDialog();
            //inputfile.Filter = "allfiles|*.xls";
            //txtInput.Text = inputfile.FileName;
            //int count = 0;
            //string[] Fname;
            //foreach(string s in inputfile.FileNames)
            //{
            //    Fname = s.Split('\\');
            //    File.Copy(s,  Fname[Fname.Length - 1]);
            //    count++;
            //}
            //MessageBox.Show(Convert.ToString(count) + " File(s) copied"); 
            #endregion
        }

        private void btninputReport_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog reportFolder = new FolderBrowserDialog();

            if(reportFolder.ShowDialog()==DialogResult.OK)
            {
                txtReport.Text = reportFolder.SelectedPath;
            }
            #region Commented Code
            //reportFolder.ShowDialog();
            //txtReport.Text = reportFolder.SelectedPath;            
            //string path = txtReport.Text;
            //StreamWriter sw = new StreamWriter(path);
            //System.IO.File.WriteAllLines(path,"Reports.csv",) 
            #endregion

        }
    }
}
