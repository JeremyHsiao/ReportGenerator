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

        // Must be updated if new report type added #NewReportType
        public enum ReportType
        {
            FullIssueDescription_TC = 0,
            FullIssueDescription_Summary,
            StandardTestReportCreation,
            KeywordIssue_Report_SingleFile,
            TC_Likely_Passed,
            FindAllKeywordInReport,
            KeywordIssue_Report_Directory,
            Excel_Sheet_Name_Update_Tool,
            FullIssueDescription_TC_report_judgement,
            TC_TestReportCreation,
            TC_AutoCorrectReport_By_Filename,
            TC_AutoCorrectReport_By_ExcelList,
            ReadAllReportHeaderIntoExcel,
            WriteAllReportHeaderAccordingToExcel,
            TC_GroupSummaryReport,
        }

        public static ReportType[] ReportSelectableTable =
        {
            ReportType.FullIssueDescription_TC,
            //ReportType.FullIssueDescription_Summary,
            //ReportType.StandardTestReportCreation,
            ReportType.KeywordIssue_Report_SingleFile,
            //ReportType.TC_Likely_Passed,
            ReportType.FindAllKeywordInReport,
            ReportType.KeywordIssue_Report_Directory,
            //ReportType.Excel_Sheet_Name_Update_Tool,
            ReportType.FullIssueDescription_TC_report_judgement,
            ReportType.TC_TestReportCreation,
            //ReportType.TC_AutoCorrectReport_By_Filename,
            ReportType.TC_AutoCorrectReport_By_ExcelList, 
            //ReportType.ReadAllReportHeaderIntoExcel,
            //ReportType.WriteAllReportHeaderAccordingToExcel, 
            //ReportType.TC_GroupSummaryReport,
         };

        //public static ReportType[] ReportSelectableTable =
        //{
        //    ReportType.FullIssueDescription_TC,
        //    ReportType.FullIssueDescription_Summary,
        //    ReportType.StandardTestReportCreation,
        //    ReportType.KeywordIssue_Report_SingleFile,
        //    ReportType.TC_Likely_Passed,
        //    ReportType.FindAllKeywordInReport,
        //    ReportType.KeywordIssue_Report_Directory,
        //    ReportType.Excel_Sheet_Name_Update_Tool,
        //    ReportType.FullIssueDescription_TC_report_judgement,
        //    ReportType.TC_TestReportCreation,
        //    ReportType.TC_AutoCorrectReport_By_Filename,
        //    ReportType.TC_AutoCorrectReport_By_ExcelList,
        //    ReportType.ReadAllReportHeaderIntoExcel,
        //    ReportType.WriteAllReportHeaderAccordingToExcel, 
        //    ReportType.TC_GroupSummaryReport,
        //};

        public static int ReportTypeToInt(ReportType type)
        {
            return (int)type;
        }

        public static ReportType ReportTypeFromInt(int int_type)
        {
            return (ReportType)int_type;
        }

        public static int ReportTypeCount = Enum.GetNames(typeof(ReportType)).Length;

        public static String GetReportName(ReportType type)
        {
            return GetReportName(ReportTypeToInt(type));
        }

        public static String GetReportName(int type_index)
        {
            return ReportName[type_index];
        }

        public static List<String> ReportNameToList()
        {
            return ReportName.ToList();
        }

        // Must be updated if new report type added #NewReportType
        private static String[] ReportName = new String[] 
        {
            "1.Issue Description for TC",
            "2.Issue Description for Summary",
            "3.Standard Test Report Creator",
            "4.Keyword Issue - Single File",
            "5.TC likely Pass",
            "6.List Keywords of all detailed reports",
            "7.Keyword Issue - Directory",
            "8.Excel sheet name update tool",
            "9.TC issue/judgement",
            "A.Jira Test Report Creator",
            "B.Auto-correct report header",
            "C.Create New Model Report",
            "D.Read All Report Header",
            "E.Write All Report Header",
            "F.Update Test Group Summary Report",
       };

        // Must be updated if new report type added #NewReportType
        private static String[][] ReportDescription = new String[][] 
        {
            // "1.Issue Description for TC",
            new String[] 
            {
                "Issue Description for TC Report", 
                "Input:",  "  Issue List + Test Case + Template (for Test case output)",
                "Output:", "  Test Case in the format of template file with linked issue in full description",
            },
            // "2.Issue Description for Summary",
            new String[] 
            {
                "Issue Description for Summary Report", 
                "Input:",  "  Issue List + Test Case + Template (for Summary Report)",
                "Output:", "  Summary in the format of template file with linked issue in full description",
            },
            // "3.Standard Test Report Creator",
            new String[] 
            {
                "Create file structure of Standard Test Report according to user's selection (Do or Not)", 
                "Input:",  "  Main Test Report File",
                "Output:", "  Directory structure and report files under directories",
            },
            // "4.Keyword Issue - Single File",
            new String[] 
            {
                "Keyword Issue to Report - Single File", 
                "Input:",  "  Test Plan/Report with Keyword",
                "Output:", "  Test Plan/Report with keyword issue list inserted on the 1st-column next to the right-side of printable area",
            },
            // "5.TC likely Pass",
            new String[] 
            {
                "Test case status is Fail but its linked issues are closed", 
                "Input:",  "  Issue List + Test Case + Template (for Test case output)",
                "Output:", "  Test Case whose linked issues are closed (other TC are hidden)",
            },
            // "6.List Keywords of all detailed reports",
            new String[] 
            {
                "Go Through all report to list down all keywords", 
                "Input:",  "  Root-directory of Report Files",
                "Output:", "  All keywords listed on output excel file",
            },
            // "7.Keyword Issue - Directory",
            new String[] 
            {
                "Keyword Issue to Report - Files under directory", 
                "Input:",  "  Test Plan/Reports with Keyword under user-specified directory",
                "Output:", "  Test Plan/Reports with keyword issue list inserted on the 1st-column next to the right-side of printable area",
            },
            // "8.Excel sheet name update tool",
            new String[] 
            {
                "Excel Reports to be checked - Files under directory", 
                "Input:",  "  Test Plan/Reports",
                "Output:", "  Test Plan/Reports with updated sheet-name if original sheet-name doesn't follow rules",
            },
            // "9.TC issue/judgement",
            new String[] 
            {
                "Issue Description + judgement from report for TC Report", 
                "Input:",  "  Issue List + Test Case + Template (for Test case output)",
                "Output:", "  Test Case in the format of template file with linked issue in full description",
            },
            // "A.Jira Test Report Creator",
            new String[] 
            {
                "Create file structure of Test Report according to TC on the Jira Test Case file", 
                "Input:",  "  Jira Test Case File & directories of source report and of output destination",
                "Output:", "  Directory structure and report files under directories",
            },
            // "B.Auto-correct report header",
            new String[] 
            {
                "Worksheet name & 1st row (header) will be auto-corrected.", 
                "Input:",  "  Root Directory of test reports",
                "Output:", "  Auto-corrected test reports",
            },
            // "C.Create New Model Report",
            new String[] 
            {
                "Worksheet name & 1st row (header) of report will be renamed and these reports are copied to corresponding folders", 
                "Input:",  "  Input excel file",
                "Output:", "  Reports copied and renamed (filename / worksheet name) according to input excel file",
            },
            // "D.Read All Report Header",
            new String[] 
            {
                "Open all Repoorts and collect header information into one Excel", 
                "Input:",  "  Report Path",
                "Output:", "  Excel file containing header value of reports under Report Path",
            },
            // "E.Write All Report Header",
            new String[] 
            {
                "Open all Repoorts and update header information according to value listed on Excel", 
                "Input:",  "  Report Path & Excel file containing header value of reports under Report Path",
                "Output:", "  Updated reports (according to header value listed on Excel",
            },
            // "F.Jira Group Summary Report Update",
            new String[] 
            {
                "Update Jira Group Summary Report (x.0)", 
                "Input:",  "  Jira Test Case File & root-directory of group summary report",
                "Output:", "  Updated reports under directories (under root-directory-plus-datetime)",
            },
        };

        private static String[][] UI_Label = new String[][] 
        {
            // "1.Issue Description for TC",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "2.Issue Description for Summary",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "3.Standard Test Report Creator",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Main Test Plan",
                "TC Template File",
            },
            // "4.Keyword Issue - Single File",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Single Test Report",
                "TC Template File",
            },
            // "5.TC likely Pass",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "6.List Keywords of all detailed reports",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "7.Keyword Issue - Directory",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "8.Excel sheet name update tool",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "9.TC issue/judgement",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "A.Jira Test Report Creator",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Source Report Path",
                "Output Report Path",
            },
            // "B.Auto-correct report header",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Source Report Path",
                "Output Report Path",
            },
            // "C.Create New Model Report",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Input Excel File",
            },
            //"D.Read All Report Header",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Output Excel File",
            },
            //"E.Write All Report Header",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Input Excel File",
            },
            // "A.Jira Test Report Creator",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Output Template",
            },
        };

        public static String GetReportDescription(int type_index)
        {
            String ret_str = "";
            ret_str += ReportGeneratorVersionString + "\r\n";
            foreach (String str in ReportDescription[type_index])
            {
                ret_str += str + "\r\n";
            }
            return ret_str;
        }

        public static String[] GetLabelTextArray(int type_index)
        {
            return UI_Label[type_index];
        }

        public static String GetReportDescription(ReportType type)
        {
            return GetReportDescription(ReportTypeToInt(type));
        }
        // END

        public void LoadConfigAll()
        {
            // config for default filename at MainForm
            this.txtBugFile.Text = XMLConfig.ReadAppSetting_String("workbook_BUG_Jira");
            this.txtTCFile.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Jira");
            this.txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
            this.txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");

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
            KeywordReport.row_test_detail_start = XMLConfig.ReadAppSetting_int("TestReport_Row_UserStart");
            KeywordReport.col_indentifier = XMLConfig.ReadAppSetting_int("TestReport_Column_Keyword_Indentifier");
            KeywordReport.col_keyword = XMLConfig.ReadAppSetting_int("TestReport_Column_Keyword_Location");
            KeywordReport.regexKeywordString = XMLConfig.ReadAppSetting_String("TestReport_Regex_Keyword_Indentifier");
            // end of config

            // config for default output directory of test report (keyword report)
            KeywordReport.TestReport_Default_Output_Dir = XMLConfig.ReadAppSetting_String("TestReport_Default_Output_Dir");

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
            Issue.D_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_D_Issue_Color");
            KeywordReport.Auto_Correct_Sheetname = XMLConfig.ReadAppSetting_Boolean("Keyword_Auto_Correct_Worksheet");

            // config for report C
            KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = XMLConfig.ReadAppSetting_Boolean("Report_C_Update_Full_Header");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = XMLConfig.ReadAppSetting_Boolean("Report_C_Update_Report_Sheetname");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = XMLConfig.ReadAppSetting_Boolean("Report_C_Clear_Keyword_Result");
            // config for header above line 21
            KeywordReport.DefaultKeywordReportHeader.Model_Name = XMLConfig.ReadAppSetting_String("Default_Model_Name");
            KeywordReport.DefaultKeywordReportHeader.Part_No = XMLConfig.ReadAppSetting_String("Default_Part_No");
            KeywordReport.DefaultKeywordReportHeader.Panel_Module = XMLConfig.ReadAppSetting_String("Default_Panel_Module");
            KeywordReport.DefaultKeywordReportHeader.TCON_Board = XMLConfig.ReadAppSetting_String("Default_TCON_Board");
            KeywordReport.DefaultKeywordReportHeader.AD_Board = XMLConfig.ReadAppSetting_String("Default_AD_Board");
            KeywordReport.DefaultKeywordReportHeader.Power_Board = XMLConfig.ReadAppSetting_String("Default_Power_Board");
            KeywordReport.DefaultKeywordReportHeader.Smart_BD_OS_Version = XMLConfig.ReadAppSetting_String("Default_Smart_BD_OS_Version");
            KeywordReport.DefaultKeywordReportHeader.Touch_Sensor = XMLConfig.ReadAppSetting_String("Default_Touch_Sensor");
            KeywordReport.DefaultKeywordReportHeader.Speaker_AQ_Version = XMLConfig.ReadAppSetting_String("Default_Speaker_AQ_Version");
            KeywordReport.DefaultKeywordReportHeader.SW_PQ_Version = XMLConfig.ReadAppSetting_String("Default_SW_PQ_Version");
            KeywordReport.DefaultKeywordReportHeader.Test_Stage = XMLConfig.ReadAppSetting_String("Default_Test_Stage");
            KeywordReport.DefaultKeywordReportHeader.Test_QTY_SN = XMLConfig.ReadAppSetting_String("Default_Test_QTY_SN");
            KeywordReport.DefaultKeywordReportHeader.Test_Period_Begin = XMLConfig.ReadAppSetting_String("Default_Test_Period_Begin");
            KeywordReport.DefaultKeywordReportHeader.Test_Period_End = XMLConfig.ReadAppSetting_String("Default_Test_Period_End");
            KeywordReport.DefaultKeywordReportHeader.Judgement = XMLConfig.ReadAppSetting_String("Default_Judgement");
            KeywordReport.DefaultKeywordReportHeader.Tested_by = XMLConfig.ReadAppSetting_String("Default_Tested_by");
            KeywordReport.DefaultKeywordReportHeader.Approved_by = XMLConfig.ReadAppSetting_String("Default_Approved_by");

            KeywordReport.DefaultKeywordReportHeader.Update_Report_Title_by_Sheetname = XMLConfig.ReadAppSetting_Boolean("Update_Report_Title_by_Sheetname");
            KeywordReport.DefaultKeywordReportHeader.Update_Model_Name = XMLConfig.ReadAppSetting_Boolean("Update_Model_Name");
            KeywordReport.DefaultKeywordReportHeader.Update_Part_No = XMLConfig.ReadAppSetting_Boolean("Update_Part_No");
            KeywordReport.DefaultKeywordReportHeader.Update_Panel_Module = XMLConfig.ReadAppSetting_Boolean("Update_Panel_Module");
            KeywordReport.DefaultKeywordReportHeader.Update_TCON_Board = XMLConfig.ReadAppSetting_Boolean("Update_TCON_Board");
            KeywordReport.DefaultKeywordReportHeader.Update_AD_Board = XMLConfig.ReadAppSetting_Boolean("Update_AD_Board");
            KeywordReport.DefaultKeywordReportHeader.Update_Power_Board = XMLConfig.ReadAppSetting_Boolean("Update_Power_Board");
            KeywordReport.DefaultKeywordReportHeader.Update_Smart_BD_OS_Version = XMLConfig.ReadAppSetting_Boolean("Update_Smart_BD_OS_Version");
            KeywordReport.DefaultKeywordReportHeader.Update_Touch_Sensor = XMLConfig.ReadAppSetting_Boolean("Update_Touch_Sensor");
            KeywordReport.DefaultKeywordReportHeader.Update_Speaker_AQ_Version = XMLConfig.ReadAppSetting_Boolean("Update_Speaker_AQ_Version");
            KeywordReport.DefaultKeywordReportHeader.Update_SW_PQ_Version = XMLConfig.ReadAppSetting_Boolean("Update_SW_PQ_Version");
            KeywordReport.DefaultKeywordReportHeader.Update_Test_Stage = XMLConfig.ReadAppSetting_Boolean("Update_Test_Stage");
            KeywordReport.DefaultKeywordReportHeader.Update_Test_QTY_SN = XMLConfig.ReadAppSetting_Boolean("Update_Test_QTY_SN");
            KeywordReport.DefaultKeywordReportHeader.Update_Test_Period_Begin = XMLConfig.ReadAppSetting_Boolean("Update_Test_Period_Begin");
            KeywordReport.DefaultKeywordReportHeader.Update_Test_Period_End = XMLConfig.ReadAppSetting_Boolean("Update_Test_Period_End");
            KeywordReport.DefaultKeywordReportHeader.Update_Judgement = XMLConfig.ReadAppSetting_Boolean("Update_Judgement");
            KeywordReport.DefaultKeywordReportHeader.Update_Tested_by = XMLConfig.ReadAppSetting_Boolean("Update_Tested_by");
            KeywordReport.DefaultKeywordReportHeader.Update_Approved_by = XMLConfig.ReadAppSetting_Boolean("Update_Approved_by");
            // end config for keyword report
        }

        private void InitializeReportFunctionListBox()
        {
            foreach (ReportType report_type in ReportSelectableTable)
            {
                String report_name;
                report_name = GetReportName((int)report_type);
                comboBoxReportSelect.Items.Add(report_name);
            }
            //int default_select_index = (int)ReportType.FullIssueDescription_Summary; // current default
            int default_select_index = 0;
            Set_comboBoxReportSelect_SelectedIndex(default_select_index);
        }

        private void Set_comboBoxReportSelect_SelectedIndex(int value)
        {
            comboBoxReportSelect.SelectedIndex = (int)ReportSelectableTable[value];
        }
        private int Get_comboBoxReportSelect_SelectedIndex()
        {
            return (int)ReportSelectableTable[comboBoxReportSelect.SelectedIndex];
        }

        static public String ReportGeneratorVersionString;

        private void MainForm_Load(object sender, EventArgs e)
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
            //this.Text = "Report Generator " + version + "   build:" + DateTime.Now.ToString("yyMMddHHmm"); ;
            //ReportGeneratorVersionString = "ReportGenerator_V" + version + DateTime.Now.ToString("(yyyyMMdd)");
            string strCompTime = Properties.Resources.BuildDate, strBuildDate = "";
            if(!string.IsNullOrEmpty(strCompTime))
            {
                string[] subs = strCompTime.Split(' ', '/'); // use ' ' & '/' as separator
                strBuildDate = "(" + subs[0]+subs[1]+subs[2] + ")";
            }
            ReportGeneratorVersionString = "ReportGenerator_V" + version + strBuildDate;
            this.Text = ReportGeneratorVersionString;

            LoadConfigAll();

            if ((Storage.GetFullPath(txtBugFile.Text) == "") ||
                (Storage.GetFullPath(txtTCFile.Text) == "") ||
                (Storage.GetFullPath(txtReportFile.Text) == "") ||
                (Storage.GetFullPath(txtOutputTemplate.Text) == ""))
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

        private bool Execute_WriteIssueDescriptionToTC(String tc_file, String template_file, String judgement_report_dir = "")
        {
            if ((ReportGenerator.global_issue_list.Count == 0) || (ReportGenerator.global_testcase_list.Count == 0) ||
                (!Storage.FileExists(tc_file)) || (!Storage.FileExists(template_file))
                || ((judgement_report_dir != "") && !Storage.DirectoryExists(judgement_report_dir)))
            {
                // protection check
                return false;
            }

            // This full issue description is needed for report purpose
            ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);

            //            ReportGenerator.WriteBacktoTCJiraExcel(tc_file);
            //ReportGenerator.WriteBacktoTCJiraExcelV2(tc_file, template_file, judgement_report_dir);
            ReportGenerator.WriteBacktoTCJiraExcelV3(tc_file, template_file, judgement_report_dir);
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

        private bool Execute_CreateTestReportbyTestCaseTask(String report_src_dir, String output_report_dir)
        {
            if (!Storage.DirectoryExists(report_src_dir) || !Storage.DirectoryExists(output_report_dir))
            {
                // protection check
                // source_dir & output_dir must exist.
                return false;
            }

            return TestReport.CopyTestReportbyTestCase(report_src_dir, output_report_dir);
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
            //List<String> report_list = Storage.FilterFilename(file_list);
            // NOTE: FilterFilename() execution is now relocated within KeywordIssueGenerationTaskV4()
            List<String> report_list = file_list;

            // This issue description is needed for report purpose
            //ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);
            ReportGenerator.global_issue_description_list_severity = Issue.GenerateIssueDescription_Severity_by_Colors(ReportGenerator.global_issue_list);
            String out_dir = KeywordReport.TestReport_Default_Output_Dir;
            if ((out_dir != "") && Storage.DirectoryExists(out_dir))
            {
                KeywordReport.KeywordIssueGenerationTaskV4(report_list, source_dir, KeywordReport.TestReport_Default_Output_Dir);
            }
            else
            {
                KeywordReport.KeywordIssueGenerationTaskV4(report_list, source_dir, Storage.GenerateDirectoryNameWithDateTime(source_dir));
            }
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

        private bool Execute_ListAllDetailedTestPlanKeywordTask(String report_root, String output_file = "")
        {
            if (!Storage.DirectoryExists(report_root))
            {
                // protection check
                return false;
            }

            List<TestPlanKeyword> keyword_list = KeywordReport.ListAllDetailedTestPlanKeywordTask(report_root, output_file);

            return true;
        }

        private bool Execute_AutoCorrectTestReportByFilename_Task(String report_root)
        {
            if (!Storage.DirectoryExists(report_root))
            {
                // protection check
                return false;
            }

            TestReport.AutoCorrectReport_by_Folder(report_root, Storage.GenerateDirectoryNameWithDateTime(report_root));

            return true;
        }

        private bool Execute_AutoCorrectTestReportByExcel_Task(String excel_input_file)
        {
            if (!Storage.FileExists(excel_input_file))
            {
                // protection check
                return false;
            }

            TestReport.AutoCorrectReport_by_Excel(excel_input_file);

            return true;
        }

        // If filename has been changed, don't change it to default at report type change afterward.
        Boolean btnSelectBugFile_Clicked = false;
        Boolean btnSelectTCFile_Clicked = false;
        Boolean btnSelectOutputTemplate_Clicked = false;
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

        // default is to select directory
        private void btnSelectReportFile_Click(object sender, EventArgs e)
        {
            int report_index = Get_comboBoxReportSelect_SelectedIndex();
            bool sel_file = false;
            String init_dir;
            switch (ReportTypeFromInt(report_index))
            {
                case ReportType.KeywordIssue_Report_SingleFile:
                    init_dir = Storage.GetFullPath(txtReportFile.Text);
                    sel_file = true;  // Here select file instead of directory
                    break;
                default:
                    init_dir = txtReportFile.Text;
                    break;
            }
            String ret_str = SelectDirectoryOrFile(init_dir, sel_file);
            if (ret_str != "")
            {
                txtReportFile.Text = ret_str;
                btnSelectReportFile_Clicked = true;
            }
        }

        private void btnSelectOutputTemplate_Click(object sender, EventArgs e)
        {
            int report_index = Get_comboBoxReportSelect_SelectedIndex();
            bool sel_file = true;
            String init_dir;
            switch (ReportTypeFromInt(report_index))
            {
                case ReportType.TC_TestReportCreation:
                    //case ReportType.FindAllKeywordInReport:
                    sel_file = false;  // Here select directory instead of file
                    init_dir = txtOutputTemplate.Text;
                    break;
                default:
                    // default is file selection here.
                    init_dir = Storage.GetFullPath(txtOutputTemplate.Text);
                    break;
            }

            String ret_str = SelectDirectoryOrFile(init_dir, sel_file);
            if (ret_str != "")
            {
                txtOutputTemplate.Text = ret_str;
                btnSelectOutputTemplate_Clicked = true;
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

            int report_index = Get_comboBoxReportSelect_SelectedIndex();

            if ((report_index < 0) || (report_index >= ReportTypeCount))
            {
                // shouldn't be out of range.
                return;
            }

            ClearIssueList();
            ClearTCList();

            UpdateUIDuringExecution(report_index: report_index, executing: true);

            MsgWindow.AppendText("Executing: " + GetReportName(report_index) + ".\n");

            Boolean open_excel_ok = ExcelAction.OpenExcelApp();
            if (open_excel_ok)
            {

                // Must be updated if new report type added #NewReportType
                switch (ReportTypeFromInt(report_index))
                {
                    case ReportType.FullIssueDescription_TC:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_WriteIssueDescriptionToTC(txtTCFile.Text, txtOutputTemplate.Text);
                        break;
                    case ReportType.FullIssueDescription_Summary:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_WriteIssueDescriptionToSummary(txtReportFile.Text);
                        break;
                    case ReportType.StandardTestReportCreation:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        bRet = Execute_CreateStandardTestReportTask(txtOutputTemplate.Text);
                        break;
                    case ReportType.KeywordIssue_Report_SingleFile:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);    // File path here
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: false);
                        break;
                    case ReportType.KeywordIssue_Report_Directory:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);     // Directory path here
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: true);
                        break;
                    case ReportType.TC_Likely_Passed:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_FindFailTCLinkedIssueAllClosed(txtTCFile.Text, txtReportFile.Text);
                        break;
                    case ReportType.FindAllKeywordInReport:
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtStandardTestReport);
                        //String main_file = txtStandardTestReport.Text;
                        //String file_dir = Storage.GetDirectoryName(main_file);
                        String output_filename = "";//use default in config file
                        String report_root_dir = Storage.GetFullPath(txtReportFile.Text);
                        bRet = Execute_ListAllDetailedTestPlanKeywordTask(report_root_dir, output_filename);
                        break;
                    case ReportType.Excel_Sheet_Name_Update_Tool:
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);     // Directory path here
                        // bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: true);
                        bRet = true;
                        break;
                    case ReportType.FullIssueDescription_TC_report_judgement:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_WriteIssueDescriptionToTC(txtTCFile.Text, txtReportFile.Text, txtOutputTemplate.Text);
                        break;
                    case ReportType.TC_TestReportCreation:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        // based on tc, create report structure
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        //String dest_dir = Storage.GetFullPath(txtReportFile.Text),
                        //       src_dir = Storage.GetFullPath(txtOutputTemplate.Text);
                        String src_dir = Storage.GetFullPath(txtReportFile.Text),
                               dest_dir = Storage.GetFullPath(txtOutputTemplate.Text);
                        bRet = Execute_CreateTestReportbyTestCaseTask(src_dir, dest_dir);
                        // update report according to jira bug
                        //if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        // to-be-implemented

                        // update judgement and header
                        // to-be-implemented
                        break;
                    case ReportType.TC_AutoCorrectReport_By_Filename:
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);
                        bRet = Execute_AutoCorrectTestReportByFilename_Task(Storage.GetFullPath(txtReportFile.Text));
                        break;
                    case ReportType.TC_AutoCorrectReport_By_ExcelList:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        // to-be-updated
                        bRet = Execute_AutoCorrectTestReportByExcel_Task(Storage.GetFullPath(txtOutputTemplate.Text));
                        break;
                    default:
                        // shouldn't be here.
                        break;
                }
            }
            else
            {
                // Open Excel application failed...
                MsgWindow.AppendText("Failed at starting Excel application.\n");
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

        private void SetEnable_ReportFile(bool value)
        {
            txtReportFile.Enabled = value;
            btnSelectReportFile.Enabled = value;
        }

        private void SetEnable_OutputTemplate(bool value)
        {
            txtOutputTemplate.Enabled = value;
            btnSelectOutputTemplate.Enabled = value;
        }

        private void comboBoxReportSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            int select = Get_comboBoxReportSelect_SelectedIndex();
            UpdateUIAfterReportTypeChanged(select);
            label_issue.Text = GetLabelTextArray(select)[0];
            label_tc.Text = GetLabelTextArray(select)[1];
            label_1st.Text = GetLabelTextArray(select)[2];
            label_2nd.Text = GetLabelTextArray(select)[3];
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
                SetEnable_ReportFile(false);
                SetEnable_OutputTemplate(false);
                btnCreateReport.Enabled = false;
            }
        }

        private void UpdateFilenameBoxUIForReportType(int ReportIndex)
        {
            // Must be updated if new report type added #NewReportType
            switch (ReportTypeFromInt(ReportIndex))
            {
                case ReportType.FullIssueDescription_TC: // "1.Issue Description for TC"
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.FullIssueDescription_Summary: // "2.Issue Description for Summary"
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.StandardTestReportCreation:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.KeywordIssue_Report_SingleFile:
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.KeywordIssue_Report_Directory:
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.TC_Likely_Passed:
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.FindAllKeywordInReport:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.Excel_Sheet_Name_Update_Tool:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.FullIssueDescription_TC_report_judgement: // "1.Issue Description for TC"
                    SetEnable_IssueFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.TC_TestReportCreation:
                    // need to rework
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.TC_AutoCorrectReport_By_Filename:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.TC_AutoCorrectReport_By_ExcelList:
                    SetEnable_IssueFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                default:
                    // Shouldn't be here
                    break;
            }
        }

        private void UpdateUIAfterReportTypeChanged(int ReportIndex)
        {
            txtReportInfo.Text = GetReportDescription(ReportIndex);
            UpdateFilenameBoxUIForReportType(ReportIndex);


            // Must be updated if new report type added #NewReportType
            switch (ReportTypeFromInt(ReportIndex))
            {
                case ReportType.FullIssueDescription_TC: // "1.Issue Description for TC"
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.FullIssueDescription_Summary: // "2.Issue Description for Summary"
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("workbook_Summary");
                    break;
                case ReportType.StandardTestReportCreation:
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_StandardTestReport");
                    break;
                case ReportType.KeywordIssue_Report_SingleFile:
                    if (!btnSelectReportFile_Clicked)  
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_Single");
                    break;
                case ReportType.KeywordIssue_Report_Directory:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    break;
                case ReportType.TC_Likely_Passed:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.FindAllKeywordInReport:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    break;
                case ReportType.Excel_Sheet_Name_Update_Tool:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = @".\SampleData\More chapters_TestCaseID";
                    break;
                case ReportType.FullIssueDescription_TC_report_judgement: // original adopted from "1.Issue Description for TC"
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.TC_TestReportCreation:
                    if (!btnSelectOutputTemplate_Clicked) // source
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("TestReport_Default_Source_Path");
                    if (!btnSelectReportFile_Clicked) // destination
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_Default_Output_Path");
                    break;
                case ReportType.TC_AutoCorrectReport_By_Filename:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    break;
                case ReportType.TC_AutoCorrectReport_By_ExcelList:
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = @".\SampleData\EVT_Winnie_Keyword2.5_keyword\Copy_Report_Excel_List.xlsx";
                    break;
                case ReportType.ReadAllReportHeaderIntoExcel:
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = @".\SampleData\EVT_Winnie_Keyword2.5_keyword\Header_Excel_List.xlsx";
                    if (!btnSelectReportFile_Clicked) // destination
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_Default_Output_Path");
                    break;
                case ReportType.WriteAllReportHeaderAccordingToExcel:
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = @".\SampleData\EVT_Winnie_Keyword2.5_keyword\Header_Excel_List.xlsx";
                    break;
                default:
                    break;
            }
        }

        private void txtTCFile_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
