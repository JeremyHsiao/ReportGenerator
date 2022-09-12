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

            // config for default filename at MainForm
            this.txtBugFile.Text = XMLConfig.ReadAppSetting("workbook_BUG_Jira");
            this.txtTCFile.Text = XMLConfig.ReadAppSetting("workbook_TC_Jira");
            this.txtReportFile.Text = XMLConfig.ReadAppSetting("workbook_Report");
            this.txtExcelTestFile.Text = XMLConfig.ReadAppSetting("workbook_ExcelTest");
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

        String [] ReportName = new String[] 
        {
            "1.Issue Description for TC",
            "2.Issue Description for Summary",
            "3.Standard Test Report Creator",
            "4.Keyword Issue to Report",
            "5.TC likely Pass"
        };

        private void InitializeReportFunctionListBox()
        {
            foreach (String name in ReportName)
            {
                comboBoxReportSelect.Items.Add(name);
            }
            comboBoxReportSelect.SelectedIndex = 0;
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
            InitializeReportFunctionListBox();
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
                ReportGenerator.global_issue_list = IssueList.GenerateIssueList(buglist_filename);
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
                ReportGenerator.global_testcase_list = TestCase.GenerateTestCaseList(tclist_filename);
                MsgWindow.AppendText("tc_list finished!\n");
                return true;
            }
        }

        private bool Execute_WriteIssueDescriptionToTC(String tc_file, String report_file)
        {
            if (ReportGenerator.global_issue_list.Count == 0)
            {
                MsgWindow.AppendText("Issue List is not available. Please read Issue list file.\n");
                return false;
            }

            if (ReportGenerator.global_testcase_list.Count == 0)
            {
                MsgWindow.AppendText("Test Case List is not available. Please read TC file.\n");
                return false;
            }

            // This full issue description is needfed for demo purpose
            ReportGenerator.global_issue_description_list = IssueList.GenerateFullIssueDescription(ReportGenerator.global_issue_list);

            // Demo 1
            ReportGenerator.WriteBacktoTCJiraExcel(tc_file);
            MsgWindow.AppendText("Writeback to tc_list finished!\n");

            return true;
        }

        private bool Execute_WriteIssueDescriptionToSummary(String tc_file, String report_file)
        {
            if (ReportGenerator.global_issue_list.Count == 0)
            {
                MsgWindow.AppendText("Issue List is not available. Please read Issue list file.\n");
                return false;
            }

            if (ReportGenerator.global_testcase_list.Count == 0)
            {
                MsgWindow.AppendText("Test Case List is not available. Please read TC file.\n");
                return false;
            }

            // Write extended string back to tc-file
            String tclist_filename = FileFunction.GetFullPath(txtTCFile.Text);
            if (!FileFunction.FileExists(tclist_filename))
            {
                return false;
            }

            String report_filename = FileFunction.GetFullPath(txtReportFile.Text);
            if (!FileFunction.FileExists(report_filename))
            {
                MsgWindow.AppendText("Report file template does not exist. Please check again.\n");
                return false;
            }

            // This full issue description is needfed for demo purpose
            ReportGenerator.global_issue_description_list = IssueList.GenerateFullIssueDescription(ReportGenerator.global_issue_list);

            // Demo 2
            ReportGenerator.SaveToReportTemplate(report_filename);
            MsgWindow.AppendText("report finished!\n");

            return true;
        }
        /*
        private void btnDemo_Click(object sender, EventArgs e)
        {
            bool bRet;

            bRet = ReadGlobalIssueListTask(txtBugFile.Text);
            if (!bRet)
            {
                return;
            }

            bRet = ReadGlobalTCListTask(txtTCFile.Text);
            if (!bRet)
            {
                return;
            }

            bRet = Execute_WriteIssueDescriptionToTC(txtTCFile.Text, txtReportFile.Text);
            if (!bRet)
            {
                return;
            }

            bRet = Execute_WriteIssueDescriptionToSummary(txtTCFile.Text, txtReportFile.Text);
            if (!bRet)
            {
                return;
            }
        }
        */
        // Because TextBox is set to Read-only, filename can be only changed via File Dialog
        // (1) no need to handle event of TestBox Text changed.
        // (2) filename (full path) is set only after File Dialog OK
        // (3) no user input --> no relative filepath --> no need to update fileanem from relative path to full path.
        private void btnSelectBugFile_Click(object sender, EventArgs e)
        {
            String init_dir = FileFunction.GetFullPath(txtBugFile.Text);
            String ret_str = FileFunction.UsesrSelectFilename(init_dir: init_dir);
            if (ret_str != "")
            {
                txtBugFile.Text = ret_str;
            }
        }

        private void btnSelectTCFile_Click(object sender, EventArgs e)
        {
            String init_dir = FileFunction.GetFullPath(txtTCFile.Text);
            String ret_str = FileFunction.UsesrSelectFilename(init_dir);
            if (ret_str != "")
            {
                txtTCFile.Text = ret_str;
            }
        }

        private void btnSelectExcelTestFile_Click(object sender, EventArgs e)
        {
            String init_dir = FileFunction.GetFullPath(txtExcelTestFile.Text);
            String ret_str = FileFunction.UsesrSelectFilename(init_dir: init_dir);
            if (ret_str != "")
            {
                txtExcelTestFile.Text = ret_str;
            }
        }

        private void btnSelectReportFile_Click(object sender, EventArgs e)
        {
            String init_dir = FileFunction.GetFullPath(txtReportFile.Text);
            String ret_str = FileFunction.UsesrSelectFilename(init_dir: init_dir);
            if (ret_str != "")
            {
                txtReportFile.Text = ret_str;
            }
        }

        private void btnGetBugList_Click(object sender, EventArgs e)
        {
            bool bRet;
            bRet = ReadGlobalIssueListTask(txtBugFile.Text);
            if (bRet)
            {
                // This full issue description is for demo purpose
                ReportGenerator.global_issue_description_list = IssueList.GenerateFullIssueDescription(ReportGenerator.global_issue_list);
            }
        }

        private void btnGetTCList_Click(object sender, EventArgs e)
        {
            bool bRet;
            bRet = ReadGlobalTCListTask(txtTCFile.Text);
        }

        // Update file path to full path (for excel operation)
        // Only enabled textbox will be updated.
        private void UpdateTextBoxPathToFull(ref TextBox box)
        {
            if (box.Enabled == false) return; // return if disabled
            String name = FileFunction.GetFullPath(box.Text);
            if (!FileFunction.FileExists(name))
            {
                return;
            }
            box.Text = name;
        }

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            bool bRet;
            
            int report_index = comboBoxReportSelect.SelectedIndex;

            if (report_index < 0)
            {
                MsgWindow.AppendText("Please select a Report.\n");
                return;
            }

            if(report_index<ReportName.Count())
            {
                MsgWindow.AppendText("Executing: " + ReportName[report_index] + ".\n");
            }

            UpdateTextBoxPathToFull(ref txtBugFile);
            UpdateTextBoxPathToFull(ref txtTCFile);
            UpdateTextBoxPathToFull(ref txtReportFile);
            UpdateTextBoxPathToFull(ref txtExcelTestFile);

            switch (comboBoxReportSelect.SelectedIndex)
            {
                 //comboBoxReportSelect.Items.Add("4.Keyword Issue to Report");
                //comboBoxReportSelect.Items.Add("5.TC likely Pass");
                case 0: 
                    //btnCreateReport.Enabled = false;
                    bRet = Execute_WriteIssueDescriptionToTC(txtTCFile.Text, txtReportFile.Text);
                    //btnCreateReport.Enabled = true;
                    break;
                case 1:
                    bRet = Execute_WriteIssueDescriptionToSummary(txtTCFile.Text, txtReportFile.Text);
                    break;
                case 2:
                    TestReport.CreateStandardTestReportTask(txtReportFile.Text);
                    break;
                case 3:
                    ReportGenerator.KeywordIssueGenerationTask(txtReportFile.Text);
                    break;
                case 4:
                    ReportGenerator.FindFailTCLinkedIssueAllClosed(txtTCFile.Text, txtReportFile.Text);
                    break;
                default:
                    break;
            }
        }

        private void btnTestExcel_Click(object sender, EventArgs e)
        {
            bool bRet;
            MsgWindow.AppendText("Start Testing Excel\n");
           // bRet = UnderDevelopment.ExcelTestMainTask(txtExcelTestFile.Text);
            MsgWindow.AppendText("Testing Excel finished!\n");
        }

        private void SetEnable_IssueFile(bool value)
        {
            txtBugFile.Enabled = value;
            btnSelectBugFile.Enabled = value;
            btnGetBugList.Enabled = value;
        }

        private void SetEnable_TCFile(bool value)
        {
            txtTCFile.Enabled = value;
            btnSelectTCFile.Enabled = value;
            btnGetTCList.Enabled = value;
        }

        private void SetEnable_Report(bool value)
        {
            txtReportFile.Enabled = value;
            btnSelectReportFile.Enabled = value;
            btnCreateReport.Enabled = value;
        }

        private void SetEnable_AdditionalFile(bool value)
        {
            txtExcelTestFile.Enabled = value;
            btnSelectExcelTestFile.Enabled = value;
            btnTestExcel.Enabled = value;
        }

        private void comboBoxReportSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBoxReportSelect.SelectedIndex)
            {
            //comboBoxReportSelect.Items.Add("3.Standard Test Report Creator");
            //comboBoxReportSelect.Items.Add("4.Keyword Issue to Report");
            //comboBoxReportSelect.Items.Add("5.TC likely Pass");
                case 0: // "1.Issue Description for TC"
                case 1: // "2.Issue Description for Summary"
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_Report(true);
                    SetEnable_AdditionalFile(false);
                   break;
                case 2:
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_Report(true);
                    SetEnable_AdditionalFile(false);
                    break;
                case 3:
                    break;
                case 4:
                    break;
                default:
                    break;
            }
        }

        private void MsgWindow_TextChanged(object sender, EventArgs e)
        {

        }

        // Update UI for difference selection.

    }
}
