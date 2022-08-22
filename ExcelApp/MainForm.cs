using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Collections.Specialized;

namespace ExcelReportApplication
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        public void LoadConfigAll()
        {
            int int_value;
            bool bool_value;

            // Read all the keys from the config file
            NameValueCollection sAll;
            sAll = ConfigurationManager.AppSettings;

            // config for default filename at MainForm
            this.txtBugFile.Text = sAll["workbook_BUG_Jira"];
            this.txtTCFile.Text = sAll["workbook_TC_Jira"];
            this.txtReportFile.Text = sAll["workbook_Report"];
            if (Boolean.TryParse(sAll["Excel_Visible"], out bool_value))
            {
                ExcelAction.ExcelVisible = bool_value;
            }

            // config for issue list
            IssueList.KeyPrefix = sAll["Issue_Key_Prefix"];
            IssueList.SheetName = sAll["Issue_SheetName"];
            if (Int32.TryParse(sAll["Issue_Row_NameDefine"], out int_value))
            {
                IssueList.NameDefinitionRow = int_value;
            }
            if (Int32.TryParse(sAll["Issue_Row_DataBegin"], out int_value))
            {
                IssueList.DataBeginRow = int_value;
            }

            // config for test-case
            TestCase.KeyPrefix = sAll["TC_Key_Prefix"];
            TestCase.SheetName = sAll["TC_SheetName"];
            if (Int32.TryParse(sAll["TC_Row_NameDefine"], out int_value))
            {
                TestCase.NameDefinitionRow = int_value;
            }
            if (Int32.TryParse(sAll["TC_Row_DataBegin"], out int_value))
            {
                TestCase.DataBeginRow = int_value;
            }

            // config for report template
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadConfigAll();

            if ((FileFunction.GetFullPath(txtBugFile.Text) == "") ||
                (FileFunction.GetFullPath(txtTCFile.Text) == "") ||
                (FileFunction.GetFullPath(txtReportFile.Text) == ""))
            {
                MsgWindow.AppendText("WARNING: one or more sample files do not exist.\n");
            }
        }

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
                ReportWorker.global_issue_list = IssueList.GenerateIssueList(buglist_filename);
                ReportWorker.global_issue_description_list = IssueList.CreateFullIssueDescription(ReportWorker.global_issue_list);
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
                ReportWorker.global_testcase_list = TestCase.GenerateTestCaseList(tclist_filename);
                MsgWindow.AppendText("tc_list finished!\n");
            }

            
            // Write extended string back to tc-file
            if (FileFunction.Exists(tclist_filename))
            {
                TestCase.WriteBacktoTCJiraExcel(tclist_filename);
                MsgWindow.AppendText("Writeback sample to tc_list finished!\n");
            }
            

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
