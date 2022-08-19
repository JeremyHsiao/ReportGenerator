using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel; 

namespace ExcelReportApplication
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            String default_data_path = @".\SampleData\";
            buglist_filename = default_data_path + workbook_BUG_Jira;
            tclist_filename = default_data_path + workbook_TC_Jira;
            report_filename = default_data_path + workbook_Report;
            txtBugFile.Text = buglist_filename;
            txtTCFile.Text = tclist_filename;
            txtReportFile.Text = report_filename;
        }

        // constant strings for workbook used in this application.
        //const string workbook_test_buglist = "SE27205_0803.xlsx";
        const string workbook_BUG_Jira = "Jira 2022-08-12T15_20_08+0800.xls";
        const string workbook_TC_Jira = "TC_Jira 2022-08-12T15_16_38+0800.xls";
        const string workbook_Report = "Report_Template.xlsx";

        private string buglist_filename, tclist_filename, report_filename, args_str;

        private void button1_Click(object sender, EventArgs e)
        {
            // Create global bug list
            // BUG_Jira
            args_str = Path.GetFullPath(buglist_filename);
            if (!File.Exists(@args_str))
            {
                MsgWindow.AppendText(args_str + " does not exist. Please check again.\n");
                return;
            }
            else
            {
                MsgWindow.AppendText("Processing bug_list:" + args_str + ".\n");
                ReportWorker.global_bug_list = ReportWorker.ProcessBugList(@args_str);
                MsgWindow.AppendText("bug_list finished!\n");
            }

            // Create global TestCase list
            args_str = Path.GetFullPath(tclist_filename);
            if (!File.Exists(@args_str))
            {
                MsgWindow.AppendText(args_str + " does not exist. Please check again.\n");
                return;
            }
            else
            {
                MsgWindow.AppendText("Processing tc_list:" + args_str + ".\n");
                ReportWorker.global_testcase_list = ReportWorker.GenerateTestCaseList(@args_str);
                MsgWindow.AppendText("tc_list finished!\n");
            }

            /*
            // Write extended string back to tc-file
            tclist_filename = workbook_TC_Jira;
            args_str = Path.GetFullPath(tclist_filename);
            if (File.Exists(@args_str))
            {
                ReportWorker.WriteBacktoTCJiraExcel(@args_str);
            }
            */

            // Write extended string to report-file (fill template and save as other file)
            args_str = Path.GetFullPath(report_filename);
            if (!File.Exists(@args_str))
            {
                MsgWindow.AppendText("Report file template does not exist. Please check again.\n");
                return;
            }
            else
            {
                ReportWorker.SaveToReportTemplate(@args_str);
                MsgWindow.AppendText("report finished!\n");
            }

            GC.Collect();
        }

        private String AskForFilename()
        {
            String ret_str = "";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = Directory.GetCurrentDirectory();
            dialog.Filter = "Excel files (*.xls/xlsx)|*.xls;*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ret_str = dialog.FileName;
            }
            return ret_str;
        }

        private void btnSelectBugFile_Click(object sender, EventArgs e)
        {
            String ret_str = AskForFilename();
            if (ret_str != "")
            {
                buglist_filename = ret_str;
                txtBugFile.Text = buglist_filename;
            }
        }

        private void btnSelectTCFile_Click(object sender, EventArgs e)
        {
            String ret_str = AskForFilename();
            if (ret_str != "")
            {
                tclist_filename = ret_str;
                txtTCFile.Text = tclist_filename;
            }
        }

        private void btnSelectReportFile_Click(object sender, EventArgs e)
        {
            String ret_str = AskForFilename();
            if (ret_str != "")
            {
                report_filename = ret_str;
                txtReportFile.Text = report_filename;
            }
        }
    }

}
