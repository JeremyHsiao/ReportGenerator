using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
            txtBugFile.Text = workbook_BUG_Jira;
            txtTCFile.Text = workbook_TC_Jira;
            txtReportFile.Text = workbook_Report;
            if ((FileFunction.GetFullPath(txtBugFile.Text) == "") ||
                (FileFunction.GetFullPath(txtTCFile.Text) == "") ||
                (FileFunction.GetFullPath(txtReportFile.Text) == ""))
            {
                MsgWindow.AppendText("WARNING: one or more sample files do not exist.\n");
            }
        }

        // constant strings for workbook used in this application.
        //const string workbook_test_buglist = "SE27205_0803.xlsx";
        const string workbook_BUG_Jira = @".\SampleData\Jira 2022-08-12T15_20_08+0800.xls";
        const string workbook_TC_Jira = @".\SampleData\TC_Jira 2022-08-12T15_16_38+0800.xls";
        const string workbook_Report = @".\SampleData\Report_Template.xlsx";

        private void button1_Click(object sender, EventArgs e)
        {
            // Create global bug list
            // BUG_Jira
            String buglist_filename = FileFunction.GetFullPath(txtBugFile.Text);
            if (!FileFunction.Exists(buglist_filename))
            {
                MsgWindow.AppendText(buglist_filename + " does not exist. Please check again.\n");
                return;
            }
            else
            {
                MsgWindow.AppendText("Processing bug_list:" + buglist_filename + ".\n");
                ReportWorker.global_bug_list = ReportWorker.ProcessBugList(buglist_filename);
                MsgWindow.AppendText("bug_list finished!\n");
            }

            // Create global TestCase list
            String tclist_filename = FileFunction.GetFullPath(txtTCFile.Text);
            if (!FileFunction.Exists(tclist_filename))
            {
                MsgWindow.AppendText(tclist_filename + " does not exist. Please check again.\n");
                return;
            }
            else
            {
                MsgWindow.AppendText("Processing tc_list:" + tclist_filename + ".\n");
                ReportWorker.global_testcase_list = ReportWorker.GenerateTestCaseList(tclist_filename);
                MsgWindow.AppendText("tc_list finished!\n");
            }

            /*
            // Write extended string back to tc-file
            if (FileFunction.Exists(tclist_filename))
            {
                ReportWorker.WriteBacktoTCJiraExcel(tclist_filename);
                MsgWindow.AppendText("Writeback sample to tc_list finished!\n");
            }
            */

            // Write extended string to report-file (fill template and save as other file)
            String report_filename = FileFunction.GetFullPath(txtReportFile.Text);
            if (!FileFunction.Exists(report_filename))
            {
                MsgWindow.AppendText("Report file template does not exist. Please check again.\n");
                return;
            }
            else
            {
                ReportWorker.SaveToReportTemplate(report_filename);
                MsgWindow.AppendText("report finished!\n");
            }

            GC.Collect();
        }

        // Because TextBox is set to Read-only, filename can be only changed via File Dialog
        // (1) no need to handle event of TestBox Text changed.
        // (2) filename (full path) is set only after File Dialog OK
        // (3) no user input --> no relative filepath --> no need to update fileanem from relative path to full path.
        private void btnSelectBugFile_Click(object sender, EventArgs e)
        {
            String ret_str = FileFunction.UsesrSelectFilename();
            if (ret_str != "")
            {
                txtBugFile.Text = ret_str;
            }
        }

        private void btnSelectTCFile_Click(object sender, EventArgs e)
        {
            String ret_str = FileFunction.UsesrSelectFilename();
            if (ret_str != "")
            {
                txtTCFile.Text = ret_str;
            }
        }

        private void btnSelectReportFile_Click(object sender, EventArgs e)
        {
            String ret_str = FileFunction.UsesrSelectFilename();
            if (ret_str != "")
            {
                txtReportFile.Text = ret_str;
            }
        }
    }

}
