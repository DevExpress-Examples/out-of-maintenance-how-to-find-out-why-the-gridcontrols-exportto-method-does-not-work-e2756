using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
                private DataTable CreateTable(int RowCount)
        {
            DataTable tbl = new DataTable();
            tbl.Columns.Add("Name", typeof(string));
            tbl.Columns.Add("ID", typeof(int));
            tbl.Columns.Add("Number", typeof(int));
            tbl.Columns.Add("Date", typeof(DateTime));
            for (int i = 0; i < RowCount; i++)
                tbl.Rows.Add(new object[] { String.Format("Name{0}", i), i, 3 - i, DateTime.Now.AddDays(i) });
            return tbl;
        }
        

        public Form1()
        {
            InitializeComponent();
            gridControl1.DataSource = CreateTable(20);
        }

        public void WriteLog(string testName, object result)
        {
            memoEdit1.Text += String.Format("[{0}] = {1}{2}", testName, result, Environment.NewLine);
        }

        public bool IsPrintingAvailable()
        {
            return gridControl1.IsPrintingAvailable;
        }

        private static string GetFileName()
        {
            FileDialog fileDialog = new SaveFileDialog();
            fileDialog.FileName = "C:\\test.xls";
            fileDialog.ShowDialog();
            string fileName = fileDialog.FileName;
            return fileName;
        }
        public object ApplicationHasRequiredRights()
        {
            object result;
            try
            {
                string fileName = String.Format("{0}EmptyFile.xls", targetFileName);
                File.Create(fileName);
                return File.Exists(fileName);

            }
            catch (Exception ex)
            {
             result = ex.Message;
            }
            return result;
        }

        public object IsFileCreated()
        {
            string fileName = String.Format("{0}GridControl.xls", targetFileName);
            gridControl1.ExportToXls(fileName);
            return File.Exists(fileName);
        }
        static string targetFileName;
        private void button1_Click(object sender, EventArgs e)
        {
            targetFileName = GetFileName();
            WriteLog("IsPrintingAvailable", IsPrintingAvailable());
            WriteLog("ApplicationHasRequiredRights", ApplicationHasRequiredRights());
            WriteLog("IsFileCreated", IsFileCreated());
        }
    }
}