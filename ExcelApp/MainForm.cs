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

            // config for default filename at MainForm
            this.txtBugFile.Text = XMLConfig.ReadAppSetting_String("workbook_BUG_Jira");
            this.txtTCFile.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Jira");
            this.txtReportFile.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
            this.txtStandardTestReport.Text = XMLConfig.ReadAppSetting_String("workbook_StandardTestReport");

            // config for default ExcelAction settings
            ExcelAction.ExcelVisible = XMLConfig.ReadAppSetting_Boolean("Excel_Visible");

            // config for issue list
            Issue.KeyPrefix = XMLConfig.ReadAppSetting_String("Issue_Key_Prefix");
            Issue.SheetName = XMLConfig.ReadAppSetting_String("Issue_SheetName");
            Issue.NameDefinitionRow = XMLConfig.ReadAppSetting_int("Issue_Row_NameDefine");
            Issue.DataBeginRow = XMLConfig.ReadAppSetting_int("Issue_Row_DataBegin");

            // config for test-case
            TestCase.KeyPrefix = XMLConfig.ReadAppSetting_String("TC_Key_Prefix");
            TestCase.SheetName = XMLConfig.ReadAppSetting_String("TC_SheetName");
            TestCase.NameDefinitionRow = XMLConfig.ReadAppSetting_int("TC_Row_NameDefine");
            TestCase.DataBeginRow = XMLConfig.ReadAppSetting_int("TC_Row_DataBegin");

            // Status string to decompose into list of string
            // Begin
            List<String> ret_list = new List<String>();
            String links = XMLConfig.ReadAppSetting_String("LinkIssueFilterStatusString");
            String[] separators = { "," };
            if ((links != null) && (links != ""))
            {
                // Separate keys into string[]
                String[] issues = links.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                if (issues != null)
                {
                    // string[] to List<String> (trimmed) and return
                    foreach (String str in issues)
                    {
                        ret_list.Add(str.Trim());
                    }
                }
            }
            ReportGenerator.fileter_status_list = ret_list;
            // End

            // config for default parameters used in Test Plan / Test Report
            TestPlan.NameDefinitionRow_TestPlan = XMLConfig.ReadAppSetting_int("TestPlan_Row_NameDefine");
            TestPlan.DataBeginRow_TestPlan = XMLConfig.ReadAppSetting_int("TestPlan_Row_DataBegin");
            TestPlan.row_test_detail_start = XMLConfig.ReadAppSetting_int("TestReport_Row_UserStart");
            TestPlan.col_indentifier = XMLConfig.ReadAppSetting_int("TestReport_Column_Keyword_Indentifier");
            TestPlan.col_keyword = XMLConfig.ReadAppSetting_int("TestReport_Column_Keyword_Location");
            TestPlan.regexKeywordString = XMLConfig.ReadAppSetting_String("TestReport_Regex_Keyword_Indentifier");
            // end of config

            // config for excel report output
            StyleString.default_font = XMLConfig.ReadAppSetting_String("default_report_Font");
            StyleString.default_size = XMLConfig.ReadAppSetting_int("default_report_FontSize");
            StyleString.default_color = XMLConfig.ReadAppSetting_Color("default_report_FontColor");
            StyleString.default_fontstyle = XMLConfig.ReadAppSetting_FontStyle("default_report_FontStyle");
            // end config for excel report output

            // config for keyword report
            Issue.A_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_A_Issue_Color");
            Issue.B_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_B_Issue_Color");
            Issue.C_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_C_Issue_Color");
            // end config for keyword report
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
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
            this.Text = "Report Generator " + version + "   build:" + DateTime.Now.ToString("yyMMddHHmm"); ;

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

        private void ClearIssueList()
        {
            ReportGenerator.global_issue_list.Clear();
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

        private void ClearTCList()
        {
            ReportGenerator.global_testcase_list.Clear();
        }

        private bool Execute_WriteIssueDescriptionToTC(String tc_file, String template_file)
        {
            if ((ReportGenerator.global_issue_list.Count == 0) || (ReportGenerator.global_testcase_list.Count == 0) ||
                (!Storage.FileExists(tc_file)) || (!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

            // This full issue description is needed for report purpose
            ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);

            //            ReportGenerator.WriteBacktoTCJiraExcel(tc_file);
            ReportGenerator.WriteBacktoTCJiraExcelV2(tc_file, template_file);
            return true;
        }

        private bool Execute_WriteIssueDescriptionToSummary(String template_file)
        {
            if ((ReportGenerator.global_issue_list.Count == 0) || (ReportGenerator.global_testcase_list.Count == 0) ||
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

        private bool Execute_KeywordIssueGenerationTask(String FileOrDirectoryName, Boolean IsDirectory = false)
        {
            List<String> file_list = new List<String>();
            String source_dir;
            if (IsDirectory == false)
            {
                if ((ReportGenerator.global_issue_list.Count == 0) || (!Storage.FileExists(FileOrDirectoryName)))
                {
                    // protection check
                    return false;
                }
                file_list.Add(FileOrDirectoryName);
                source_dir = Storage.GetDirectoryName(FileOrDirectoryName);
            }
            else
            {
                if ((ReportGenerator.global_issue_list.Count == 0) || (!Storage.DirectoryExists(FileOrDirectoryName)))
                {
                    // protection check
                    return false;
                }
                file_list = Storage.ListFilesUnderDirectory(FileOrDirectoryName);
                //MsgWindow.AppendText("File found under directory " + FileOrDirectoryName + "\n");
                //foreach (String filename in file_list)
                //    MsgWindow.AppendText(filename + "\n");
                source_dir = FileOrDirectoryName;
            }
            // filename check to exclude non-report files.
            List<String> report_list = Storage.FilterFilename(file_list);

            // This issue description is needed for report purpose
            //ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);
            ReportGenerator.global_issue_description_list_severity = Issue.GenerateIssueDescription_Severity_by_Colors(ReportGenerator.global_issue_list);
            KeywordReport.KeywordIssueGenerationTaskV4(report_list, source_dir, Storage.GenerateDirectoryNameWithDateTime(source_dir));
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
            if (!Storage.FileExists(main_file) || !Storage.DirectoryExists(report_root))
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
                ClearIssueList();
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
                ClearTCList();
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
                case ReportGenerator.ReportType.KeywordIssue_Report_Directory:
                    //case ReportGenerator.ReportType.FindAllKeywordInReport:
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
                MsgWindow.AppendText(box.Text + "can't be found. Please check again.\n");
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

            if ((report_index < 0) || (report_index >= ReportGenerator.ReportTypeCount))
            {
                // shouldn't be out of range.
                return;
            }

            ClearIssueList();
            ClearTCList();

            UpdateUIDuringExecution(report_index: report_index, executing: true);

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
                case ReportGenerator.ReportType.KeywordIssue_Report_SingleFile:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                    UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);    // File path here
                    if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                    bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: false);
                    break;
                case ReportGenerator.ReportType.KeywordIssue_Report_Directory:
                    UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                    UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);     // Directory path here
                    if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                    bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: true);
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
                case ReportGenerator.ReportType.Excel_Sheet_Name_Update_Tool:
                    UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);     // Directory path here
                    // bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: true);
                    bRet = true;
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
                case ReportGenerator.ReportType.KeywordIssue_Report_SingleFile:
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(false);
                    SetEnable_OutputFile(true);
                    SetEnable_StandardReport(false);
                    break;
                case ReportGenerator.ReportType.KeywordIssue_Report_Directory:
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
                case ReportGenerator.ReportType.Excel_Sheet_Name_Update_Tool:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_OutputFile(true);
                    SetEnable_StandardReport(false);
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
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportGenerator.ReportType.FullIssueDescription_Summary: // "2.Issue Description for Summary"
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("workbook_Summary");
                    break;
                case ReportGenerator.ReportType.StandardTestReportCreation:
                    if (!btnSelectExcelTestFile_Clicked)
                        txtStandardTestReport.Text = XMLConfig.ReadAppSetting_String("workbook_StandardTestReport");
                    break;
                case ReportGenerator.ReportType.KeywordIssue_Report_SingleFile:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = @".\SampleData\A.1.1_OSD _All.xlsx";
                    break;
                case ReportGenerator.ReportType.KeywordIssue_Report_Directory:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = @".\SampleData\More chapters_TestCaseID";
                    break;
                case ReportGenerator.ReportType.TC_Likely_Passed:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportGenerator.ReportType.FindAllKeywordInReport:
                    if (!btnSelectExcelTestFile_Clicked)
                        txtStandardTestReport.Text = XMLConfig.ReadAppSetting_String("workbook_ReportToTestKeyword");
                    break;
                case ReportGenerator.ReportType.Excel_Sheet_Name_Update_Tool:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = @".\SampleData\More chapters_TestCaseID";
                    break;
                default:
                    break;
            }
        }
    }
}
