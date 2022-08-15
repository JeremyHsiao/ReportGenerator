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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        // constant strings for workbook used in this application.
        //const string workbook_test_buglist = "SE27205_0803.xlsx";
        const string workbook_BUG_Jira = "Jira 2022-08-12T15_20_08+0800.xls";
        const string workbook_TC_Jira = "TC_Jira.xls";

        private Dictionary<string, List<StyleString>> global_bug_list = new Dictionary<string, List<StyleString>>();

        private void button1_Click(object sender, EventArgs e)
        {
            string buglist_filename, tclist_filename, args_str;

            // BUG_Jira
            buglist_filename = workbook_BUG_Jira;
            args_str = Directory.GetCurrentDirectory() + '\\' + buglist_filename;
            if (!File.Exists(@args_str))
            {
                MsgWindow.AppendText(args_str + " does not exist. Please check again.\n");
            }
            else
            {
                MsgWindow.AppendText("Processing bug_list:" + args_str + ".\n");
                global_bug_list = ReportWorker.ProcessJiraBugFile(@args_str);
                MsgWindow.AppendText("bug_list finished!\n");
            }

            tclist_filename = workbook_TC_Jira;
            args_str = Directory.GetCurrentDirectory() + '\\' + tclist_filename;
            if (!File.Exists(@args_str))
            {
                MsgWindow.AppendText(args_str + " does not exist. Please check again.\n");
            }
            else
            {
                MsgWindow.AppendText("Processing tc_list:" + args_str + ".\n");
                ReportWorker.ProcessTCJiraExcel(@args_str, global_bug_list);
                MsgWindow.AppendText("tc_list finished!\n");
            }
            GC.Collect();
        }
    }

}
