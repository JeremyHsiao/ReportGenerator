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
        }

        // constant strings for workbook used in this application.
        //const string workbook_test_buglist = "SE27205_0803.xlsx";
        const string workbook_BUG_Jira = "Jira 2022-08-12T15_20_08+0800.xls";
        const string workbook_TC_Jira = "TC_Jira 2022-08-12T15_16_38+0800.xls";
        const string workbook_Report = "Report_Template.xlsx";

        private void button1_Click(object sender, EventArgs e)
        {
            string buglist_filename, tclist_filename, report_filename, args_str;

            // Create global bug list
            // BUG_Jira
            buglist_filename = workbook_BUG_Jira;
            args_str = Directory.GetCurrentDirectory() + '\\' + buglist_filename;
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
            tclist_filename = workbook_TC_Jira;
            args_str = Directory.GetCurrentDirectory() + '\\' + tclist_filename;
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
            args_str = Directory.GetCurrentDirectory() + '\\' + tclist_filename;
            if (File.Exists(@args_str))
            {
                ReportWorker.WriteBacktoTCJiraExcel(@args_str);
            }
            */

            // Write extended string to report-file (fill template and save as other file)
            report_filename = workbook_Report;
            args_str = Directory.GetCurrentDirectory() + '\\' + report_filename;
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
    }

}
