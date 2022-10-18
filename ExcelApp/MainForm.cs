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
            this.txtReportFile.Text = XMLConfig.ReadAppSetting("workbook_TC_Template");
            this.txtStandardTestReport.Text = XMLConfig.ReadAppSetting("workbook_StandardTestReport");
            if (Boolean.TryParse(XMLConfig.ReadAppSetting("Excel_Visible"), out bool_value))
            {
                ExcelAction.ExcelVisible = bool_value;
            }

            // config for issue list
            Issue.KeyPrefix = XMLConfig.ReadAppSetting("Issue_Key_Prefix");
            Issue.SheetName = XMLConfig.ReadAppSetting("Issue_SheetName");
            if (Int32.TryParse(XMLConfig.ReadAppSetting("Issue_Row_NameDefine"), out int_value))
            {
                Issue.NameDefinitionRow = int_value;
            }
            if (Int32.TryParse(XMLConfig.ReadAppSetting("Issue_Row_DataBegin"), out int_value))
            {
                Issue.DataBeginRow = int_value;
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

        private void InitializeReportFunctionListBox()
        {

            foreach (String name in ReportGenerator.ReportNameToList())
            {
                comboBoxReportSelect.Items.Add(name);
            }
            comboBoxReportSelect.SelectedIndex = 0; // (int)ReportGenerator.ReportType.FullIssueDescription_Summary; // current default
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadConfigAll();

            if ((Storage.GetFullPath(txtBugFile.Text) == "") ||
                (Storage.GetFullPath(txtTCFile.Text) == "") ||
                (Storage.GetFullPath(txtReportFile.Text) == "") ||
                (Storage.GetFullPath(txtStandardTestReport.Text) == ""))
            {
                MsgWindow.AppendText("WARNING: one or more sample files do not exist.\n");
            }
            InitializeReportFunctionListBox();
        }

        private bool ReadGlobalIssueListTask(String filename)
        {
            String buglist_filename = Storage.GetFullPath(filename);
            if (!Storage.FileExists(buglist_filename))
            {
                MsgWindow.AppendText(buglist_filename + " does not exist. Please check again.\n");
                return false;
            }
            else
            {
                MsgWindow.AppendText("Processing bug_list:" + buglist_filename + ".\n");
                ReportGenerator.global_issue_list = Issue.GenerateIssueList(buglist_filename);
                MsgWindow.AppendText("bug_list finished!\n");
                return true;
            }
        }

        private bool ReadGlobalTCListTask(String filename)
        {
            String tclist_filename = Storage.GetFullPath(filename);
            if (!Storage.FileExists(tclist_filename))
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

        private bool LoadIssueListIfEmpty(String filename)
        {
            if (ReportGenerator.global_issue_list.Count == 0)
            {
                return ReadGlobalIssueListTask(filename);
            }
            else
            {
                return true;
            }
        }

        private bool LoadTCListIfEmpty(String filename)
        {
            if (ReportGenerator.global_testcase_list.Count == 0)
            {
                return ReadGlobalTCListTask(filename);
            }
            else
            {
                return true;
            }
        }

        private bool Execute_WriteIssueDescriptionToTC(String tc_file, String template_file)
        {
            if ((ReportGenerator.global_issue_list.Count == 0)||(ReportGenerator.global_testcase_list.Count == 0)||
                (!Storage.FileExists(tc_file))||(!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

            // This full issue description is needed for report purpose
            ReportGenerator.global_full_issue_description_list = Issue.GenerateFullIssueDescription(ReportGenerator.global_issue_list);

//            ReportGenerator.WriteBacktoTCJiraExcel(tc_file);
            ReportGenerator.WriteBacktoTCJiraExcelV2(tc_file, template_file);
            return true;
        }

        private bool Execute_WriteIssueDescriptionToSummary(String template_file)
        {
            if ((ReportGenerator.global_issue_list.Count == 0)||(ReportGenerator.global_testcase_list.Count == 0)||
                (!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

            // This full issue description is needed for report purpose
            ReportGenerator.global_full_issue_description_list = Issue.GenerateFullIssueDescription(ReportGenerator.global_issue_list);

            SummaryReport.SaveIssueToSummaryReport(template_file);

            return true;
        }

        private bool Execute_CreateStandardTestReportTask(String main_file)
        {
            if (!Storage.FileExists(main_file))
            {
                // protection check
                return false;
            }

            return TestReport.CreateStandardTestReportTask(main_file);
        }

        private bool Execute_KeywordIssueGenerationTask(String template_file)
        {
            if ((ReportGenerator.global_issue_list.Count == 0) || (!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

            // This issue description is needed for report purpose
            ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);

            KeywordReport.KeywordIssueGenerationTaskV2(template_file);
            return true;
        }

        private bool Execute_FindFailTCLinkedIssueAllClosed(String tc_file, String template_file)
        {
            if ((ReportGenerator.global_issue_list.Count == 0) || (ReportGenerator.global_testcase_list.Count == 0) ||
                (!Storage.FileExists(tc_file)) || (!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

            // This issue description is needed for report purpose
            ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);

            ReportGenerator.FindFailTCLinkedIssueAllClosed(tc_file, template_file);
            return true;
        }

        private bool Execute_ListAllDetailedTestPlanKeywordTask(String main_file, String report_root)
        {
            if (!Storage.FileExists(main_file)||!Storage.DirectoryExists(report_root))
            {
                // protection check
                return false;
            }

            List<TestPlanKeyword> keyword_list = KeywordReport.ListAllDetailedTestPlanKeywordTask(main_file, report_root);

            MsgWindow.AppendText("Keyword list:\n");
            MsgWindow.AppendText("-------------\n");
            foreach (TestPlanKeyword keyword in keyword_list)
            {
                MsgWindow.AppendText(keyword.Keyword + " @ " + keyword.Worksheet + "(" + keyword.AtRow + ")\n");
            }
            MsgWindow.AppendText("-------------\n");
            return true;
        }
        
        // If filename has been changed, don't change it to default at report type change afterward.
        Boolean btnSelectBugFile_Clicked = false;
        Boolean btnSelectTCFile_Clicked = false;
        Boolean btnSelectExcelTestFile_Clicked = false;
        Boolean btnSelectReportFile_Clicked = false;
        
        // Because TextBox is set to Read-only, filename can be only changed via File Dialog
        // (1) no need to handle event of TestBox Text changed.
        // (2) filename (full path) is set only after File Dialog OK
        // (3) no user input --> no relative filepath --> no need to update fileanem from relative path to full path.
        private void btnSelectBugFile_Click(object sender, EventArgs e)
        {
            String init_dir = Storage.GetFullPath(txtBugFile.Text);
            String ret_str = Storage.UsesrSelectFilename(init_dir: init_dir);
            if (ret_str != "")
            {
                txtBugFile.Text = ret_str;
                btnSelectBugFile_Clicked = true;
            }
        }

        private void btnSelectTCFile_Click(object sender, EventArgs e)
        {
            String init_dir = Storage.GetFullPath(txtTCFile.Text);
            String ret_str = Storage.UsesrSelectFilename(init_dir);
            if (ret_str != "")
            {
                txtTCFile.Text = ret_str;
                btnSelectTCFile_Clicked = true;
            }
        }

        private void btnSelectExcelTestFile_Click(object sender, EventArgs e)
        {
            String init_dir = Storage.GetFullPath(txtStandardTestReport.Text);
            String ret_str = Storage.UsesrSelectFilename(init_dir: init_dir);
            if (ret_str != "")
            {
                txtStandardTestReport.Text = ret_str;
                btnSelectExcelTestFile_Clicked = true;
            }
        }

        private void btnSelectReportFile_Click(object sender, EventArgs e)
        {
            int report_index = comboBoxReportSelect.SelectedIndex;
            bool sel_file = true;
            switch (ReportGenerator.ReportTypeFromInt(report_index))
            {
                case ReportGenerator.ReportType.FindAllKeywordInReport:
                    sel_file = false;  // Here select directory instead of file
                    break;
            }

            String init_dir = Storage.GetFullPath(btnSelectReportFile.Text);
            String ret_str = SelectDirectoryOrFile(init_dir, sel_file);
            if (ret_str != "")
            {
                txtReportFile.Text = ret_str;
                btnSelectReportFile_Clicked = true;
            }
        }

        private String SelectDirectoryOrFile(String init_text, bool sel_file = true)
        {
            String init_dir = Storage.GetFullPath(init_text), ret_str;
            if (sel_file == true)
            {
                ret_str = Storage.UsesrSelectFilename(init_dir: init_dir);
            }
            else
            {
                ret_str = Storage.UsersSelectDirectory(init_dir: init_dir);
            }
            return ret_str;
        }
       
/*
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
*/
        // Update file path to full path (for excel operation)
        // Only enabled textbox will be updated.
        private void UpdateTextBoxPathToFullAndCheckExist(ref TextBox box)
        {
            String name = Storage.GetFullPath(box.Text);
            if (!Storage.FileExists(name))
            {
                MsgWindow.AppendText( box.Text + "can't be found. Please check again.\n");
                return;
            }
            box.Text = name;
        }

        private void UpdateTextBoxDirToFullAndCheckExist(ref TextBox box)
        {
            String name = Storage.GetFullPath(box.Text);
            if (!Storage.DirectoryExists(name))
            {
                MsgWindow.AppendText(box.Text + "can't be found. Please check again.\n");
                return;
            }
            box.Text = name;
        }

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            bool bRet;
            
            int report_index = comboBoxReportSelect.SelectedIndex;

            if ((report_index < 0)||(report_index >= ReportGenerator.ReportTypeCount))
            {
                // shouldn't be out of range.
                return;
            }

            UpdateUIDuringExecution ( report_index: report_index, executing: true);

            MsgWindow.AppendText("Executing: " + ReportGenerator.GetReportName(report_index) + ".\n");

            ExcelAction.OpenExcelApp();

            // Must be updated if new report type added #NewReportType
            switch (ReportGenerator.ReportTypeFromInt(report_index))
            {
                case ReportGenerator.ReportType.FullIssueDescription_TC:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                    if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                    if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                    bRet = Execute_WriteIssueDescriptionToTC(txtTCFile.Text, txtReportFile.Text);
                    break;
                case ReportGenerator.ReportType.FullIssueDescription_Summary:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                    if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                    if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                    bRet = Execute_WriteIssueDescriptionToSummary(txtReportFile.Text);
                    break;
                case ReportGenerator.ReportType.StandardTestReportCreation:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtStandardTestReport);
                    bRet = Execute_CreateStandardTestReportTask(txtStandardTestReport.Text);
                    break;
                case ReportGenerator.ReportType.KeywordIssue_Report:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                    if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                    bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text);
                    break;
                case ReportGenerator.ReportType.TC_Likely_Passed:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                    if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                    if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                    bRet = Execute_FindFailTCLinkedIssueAllClosed(txtTCFile.Text, txtReportFile.Text);
                    break;
                case ReportGenerator.ReportType.FindAllKeywordInReport:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtStandardTestReport);
                    String main_file = txtStandardTestReport.Text;
                    String file_dir = Storage.GetDirectoryName(main_file);
                    String report_root_dir = Storage.GetDirectoryName(file_dir);
                    bRet = Execute_ListAllDetailedTestPlanKeywordTask(main_file, report_root_dir);
                    break;
                default:
                    // shouldn't be here.
                    break;
            }

            ExcelAction.CloseExcelApp();

            MsgWindow.AppendText("Finished.\n");
            UpdateUIDuringExecution(report_index: report_index, executing: false);
        }

        private void SetEnable_IssueFile(bool value)
        {
            txtBugFile.Enabled = value;
            btnSelectBugFile.Enabled = value;
        }

        private void SetEnable_TCFile(bool value)
        {
            txtTCFile.Enabled = value;
            btnSelectTCFile.Enabled = value;
        }

        private void SetEnable_OutputFile(bool value)
        {
            txtReportFile.Enabled = value;
            btnSelectReportFile.Enabled = value;
        }

        private void SetEnable_StandardReport(bool value)
        {
            txtStandardTestReport.Enabled = value;
            btnSelectExcelTestFile.Enabled = value;
        }

        private void comboBoxReportSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateUIAfterReportTypeChanged(comboBoxReportSelect.SelectedIndex);
        }

        private void UpdateUIDuringExecution(int report_index, bool executing)
        {
            if (!executing)
            {
                UpdateFilenameBoxUIForReportType(report_index);
                btnCreateReport.Enabled = true;
            }
            else
            {
                SetEnable_IssueFile(false);
                SetEnable_TCFile(false);
                SetEnable_OutputFile(false);
                SetEnable_StandardReport(false);
                btnCreateReport.Enabled = false;
            }
        }

        private void UpdateFilenameBoxUIForReportType(int ReportIndex)
        {
            // Must be updated if new report type added #NewReportType
            switch (ReportGenerator.ReportTypeFromInt(ReportIndex))
            {
                case ReportGenerator.ReportType.FullIssueDescription_TC: // "1.Issue Description for TC"
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_OutputFile(true);
                    SetEnable_StandardReport(false);
                     break;
                case ReportGenerator.ReportType.FullIssueDescription_Summary: // "2.Issue Description for Summary"
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_OutputFile(true);
                    SetEnable_StandardReport(false);
                    break;
                case ReportGenerator.ReportType.StandardTestReportCreation:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_OutputFile(false);
                    SetEnable_StandardReport(true);
                    break;
                case ReportGenerator.ReportType.KeywordIssue_Report:
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(false);
                    SetEnable_OutputFile(true);
                    SetEnable_StandardReport(false);
                    break;
                case ReportGenerator.ReportType.TC_Likely_Passed:
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_OutputFile(true);
                    SetEnable_StandardReport(false);
                    break;
                case ReportGenerator.ReportType.FindAllKeywordInReport:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_OutputFile(false);
                    SetEnable_StandardReport(true);
                    break;
                default:
                    // Shouldn't be here
                    break;
            }
        }

        private void UpdateUIAfterReportTypeChanged(int ReportIndex)
        {
            txtReportInfo.Text = ReportGenerator.GetReportDescription(ReportIndex);
            UpdateFilenameBoxUIForReportType(ReportIndex);
            // Must be updated if new report type added #NewReportType
            switch (ReportGenerator.ReportTypeFromInt(ReportIndex))
            {
                case ReportGenerator.ReportType.FullIssueDescription_TC: // "1.Issue Description for TC"
                    if(!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting("workbook_TC_Template");
                    break;
                case ReportGenerator.ReportType.FullIssueDescription_Summary: // "2.Issue Description for Summary"
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting("workbook_Summary");
                    break;
                case ReportGenerator.ReportType.StandardTestReportCreation:
                    if (!btnSelectExcelTestFile_Clicked)
                        txtStandardTestReport.Text = XMLConfig.ReadAppSetting("workbook_StandardTestReport");
                    break;
                case ReportGenerator.ReportType.KeywordIssue_Report:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = @".\SampleData\A.1.1_OSD _All.xlsx";
                    break;
                case ReportGenerator.ReportType.TC_Likely_Passed:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting("workbook_TC_Template");
                    break;
                case ReportGenerator.ReportType.FindAllKeywordInReport:
                    if (!btnSelectExcelTestFile_Clicked)
                    txtStandardTestReport.Text = XMLConfig.ReadAppSetting("workbook_ReportToTestKeyword");
                    break;
                default:
                    break;
            }
        }
    }
}
