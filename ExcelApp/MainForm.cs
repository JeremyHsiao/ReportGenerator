﻿using System;
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

            // config for default filename at MainForm
            this.txtBugFile.Text = XMLConfig.ReadAppSetting("workbook_BUG_Jira");
            this.txtTCFile.Text = XMLConfig.ReadAppSetting("workbook_TC_Jira");
            this.txtReportFile.Text = XMLConfig.ReadAppSetting("workbook_Report");

            if (Boolean.TryParse(XMLConfig.ReadAppSetting("Excel_Visible"), out bool_value))
            {
                ExcelAction.ExcelVisible = bool_value;
            }

            // config for issue list
            IssueList.KeyPrefix = XMLConfig.ReadAppSetting("Issue_Key_Prefix");
            IssueList.SheetName = XMLConfig.ReadAppSetting("Issue_SheetName");
            if (Int32.TryParse(XMLConfig.ReadAppSetting("Issue_Row_NameDefine"), out int_value))
            {
                IssueList.NameDefinitionRow = int_value;
            }
            if (Int32.TryParse(XMLConfig.ReadAppSetting("Issue_Row_DataBegin"), out int_value))
            {
                IssueList.DataBeginRow = int_value;
            }

            // config for test-case
            TestCase.KeyPrefix = XMLConfig.ReadAppSetting("TC_Key_Prefix");
            TestCase.SheetName = XMLConfig.ReadAppSetting("TC_SheetName");
            if (Int32.TryParse(XMLConfig.ReadAppSetting("TC_Row_NameDefine"), out int_value))
            {
                TestCase.NameDefinitionRow = int_value;
            }
            if (Int32.TryParse(XMLConfig.ReadAppSetting("TC_Row_DataBegin"), out int_value))
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

        private bool ReadGlobalIssueListTask(String filename)
        {
            String buglist_filename = FileFunction.GetFullPath(filename);
            if (!FileFunction.FileExists(buglist_filename))
            {
                MsgWindow.AppendText(buglist_filename + " does not exist. Please check again.\n");
                return false;
            }
            else
            {
                MsgWindow.AppendText("Processing bug_list:" + buglist_filename + ".\n");
                KeywordIssue.global_issue_list = IssueList.GenerateIssueList(buglist_filename);
                MsgWindow.AppendText("bug_list finished!\n");
                return true;
            }
        }

        private bool ReadGlobalTCListTask(String filename)
        {
            String tclist_filename = FileFunction.GetFullPath(filename);
            if (!FileFunction.FileExists(tclist_filename))
            {
                MsgWindow.AppendText(tclist_filename + " does not exist. Please check again.\n");
                return false;
            }
            else
            {
                MsgWindow.AppendText("Processing tc_list:" + tclist_filename + ".\n");
                KeywordIssue.global_testcase_list = TestCase.GenerateTestCaseList(tclist_filename);
                MsgWindow.AppendText("tc_list finished!\n");
                return true;
            }
        }

        private bool SaveKeywordIssueTask(String tc_file, String report_file)
        {
            if (KeywordIssue.global_issue_list.Count == 0)
            {
                bool bRet = ReadGlobalIssueListTask(txtBugFile.Text);
                if (!bRet)
                {
                    //MsgWindow.AppendText("Issue List not available. Please check Issue list file.\n");
                    return false;
                }
            }

            if (KeywordIssue.global_testcase_list.Count == 0)
            {
                bool bRet = ReadGlobalTCListTask(txtTCFile.Text);
                if (!bRet)
                {
                    //MsgWindow.AppendText("Test Case List is not available. Please check TC file.\n");
                    return false;
                }
            }

            String report_filename = FileFunction.GetFullPath(txtReportFile.Text);
            if (!FileFunction.FileExists(report_filename))
            {
                MsgWindow.AppendText("Report file does not exist. Please check again.\n");
                return false;
            }

            // This full issue description is needfed for keyword issue list
            KeywordIssue.global_issue_description_list = IssueList.GenerateIssueSummary(KeywordIssue.global_issue_list);

            // Read report file for keyword & its row and store into keyword/row dictionary
            Dictionary<String, int> keyword_row = new Dictionary<String, int>();

            // Generate keyword issue list and store into keyword/issue_summary_list_style_string dictionary
            Dictionary<String, List<StyleString>> keyword_issue = new Dictionary<String, List<StyleString>>();

            // Write keyword issue list back to report file and Save.
            KeywordIssue.SaveToReportTemplate(report_filename); // to be updated

            //
            MsgWindow.AppendText("report finished!\n");

            return true;
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

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            bool bRet;
            bRet = SaveKeywordIssueTask(txtTCFile.Text, txtReportFile.Text);
        }

    }
}
