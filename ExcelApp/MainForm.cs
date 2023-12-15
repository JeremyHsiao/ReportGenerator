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
using System.Management;
using System.Threading;

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
            CreateImportToJiraCSV,
            KeywordIssue_Report_SingleFile,
            TC_Likely_Passed,
            FindAllKeywordInReport,
            KeywordIssue_Report_Directory,                  // Report 7
            Excel_Sheet_Name_Update_Tool,
            FullIssueDescription_TC_report_judgement,       // Report 9
            CreateCSTReport,                                // Report A -- Create CST
            TC_AutoCorrectReport_By_Filename,
            TC_AutoCorrectReport_By_ExcelList,              // Report C
            CopyReportOnly,                                 // Report D -- copy only version of report C
            RemoveInternalSheet,                            // Report E -- remove internalsheet version of report C
            TC_GroupSummaryReport,                          // Report F -- Not used
            Update_Report_Linked_Issue,
            Update_Keyword_and_TC_Report,                   // Report H -- it is report 7 + 9
            Man_Power_Processing,                           // Report I -- man-power
            Update_Repoart_A_then_Report_H,                 // Report J -- it is report A +Ｈ
            Update_Report_Linked_Issue_and_TC_Report,       // Report K -- simplified version of report H - no keyword at all, less format update/correction
        }

        public static ReportType[] ReportSelectableTable =
        {
            ReportType.FullIssueDescription_TC,                     // report 1
            //ReportType.FullIssueDescription_Summary,
            //ReportType.StandardTestReportCreation,
            ReportType.KeywordIssue_Report_SingleFile,
            //ReportType.TC_Likely_Passed,
            ReportType.FindAllKeywordInReport,
            ReportType.KeywordIssue_Report_Directory,               // report 7
            //ReportType.Excel_Sheet_Name_Update_Tool,
            ReportType.FullIssueDescription_TC_report_judgement,    // report 9
            ReportType.CreateCSTReport,                             // report A
            //ReportType.TC_AutoCorrectReport_By_Filename,
            ReportType.TC_AutoCorrectReport_By_ExcelList, 
            ReportType.CopyReportOnly,
            ReportType.RemoveInternalSheet, 
            //ReportType.TC_GroupSummaryReport,
            //ReportType.Update_Report_Linked_Issue,
            ReportType.Update_Keyword_and_TC_Report,                // Report H
            //ReportType.Man_Power_Processing,
            //ReportType.Update_Repoart_A_then_Report_H,
            ReportType.Update_Report_Linked_Issue_and_TC_Report,    // Report K
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
        //    ReportType.CopyReportOnly,
        //    ReportType.RemoveInternalSheet, 
        //    ReportType.TC_GroupSummaryReport,
        //    ReportType.Update_Report_Linked_Issue,
        //    ReportType.Update_Keyword_and_TC_Report,
        //    ReportType.Man_Power_Processing,
        //    ReportType.Update_Repoart_A_then_Report_H,
        //};

        public static int ReportTypeCount = Enum.GetNames(typeof(ReportType)).Length;

        public static String GetReportName(ReportType report_type)
        {
            int report_type_index = (int)(report_type);
            // prevent out of boundary
            if (report_type_index < ReportTypeCount)
            {
                return ReportName[report_type_index];
            }
            else
            {
                return "GetReportName_issue";
            }
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
            "3.Create ImportToJira CSV",
            "4.Keyword Issue - Single File",
            "5.TC likely Pass",
            "6.List Keywords of all detailed reports",
            "7.Keyword Issue - Directory",
            "8.Excel sheet name update tool",
            "9.TC issue/judgement",
            "A.Create CST Report",
            "B.Auto-correct report header",
            "C.Create New Model Report",
            "D.Copy Report Only",
            "E.Remove Internal Sheets from Report",
            "F.Update Test Group Summary Report",
            "G.Update Report Linked Issue",
            "H.Update Keyword Report and TC summary (7+9)",
            "I.Man-Power Processing",
            "J.Create Report, Update Report and TC Summary (A+H)",
            "K.Update Report and TC summary (linked issue)",
            // out-of-boundary
            "Z.Final Report",
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
                "Create CSV file containing data field to import Jira testcase to TC project", 
                "Input:",  "  Excel file contains testplan, ImportToJira Template, testcase and report list (for assignee)",
                "Output:", "  ImportToJira CSV file",
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
                "Create CST Report", 
                "Input:",  "  Input excel file",
                "Output:", "  Reports copied and renamed (filename / worksheet name) according to input excel file",
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
                "Create New Model Report", 
                "Input:",  "  Input excel file",
                "Output:", "  Reports copied and renamed (filename / worksheet name) according to input excel file",
            },
            // "D.Copy Report Only",
            new String[] 
            {
                "Reports copied to destination. Copy files only so that contents are not touched", 
                "Input:",  "  Report Path",
                "Output:", "  Reports copied to destination.",
            },
            // "E.Remove Internal Sheets from Report",
            new String[] 
            {
                "Reports' internal sheets are removed and saved to destination.", 
                "Input:",  "  Report Path",
                "Output:", "  Reports saved to destination.",
            },
            // "F.Update Test Group Summary Report",
            new String[] 
            {
                "Update Jira Group Summary Report (x.0)", 
                "Input:",  "  Jira Test Case File & root-directory of group summary report",
                "Output:", "  Updated reports under directories (under root-directory-plus-datetime)",
            },
            // "G.Update Report Linked Issue",
            new String[] 
            {
                "Update Linked Issue field in each Report", 
                "Input:",  "  Jira Bug & TC file, root-directory of reports to be updated",
                "Output:", "  Updated reports under directories (named by root-directory-plus-datetime)",
            },
            // "H.Update Keyword Rerpot and TC summary (7+9)",
            new String[] 
            {
                "Update keyword report and TC Linked Issue", 
                "Input:",  "  Jira Bug & TC file, Template (for Test case output), and root-directory of reports to be updated",
                "Output:", "  Updated reports under directories (named by root-directory-plus-datetime) and TC summary with Linked issues",
            },
            // "I.Man-Power Processing",
            new String[] 
            {
                "Augment CSV exported from Jira man-power plan with information required for Pivot Table", 
                "Input:",  "  CSV exported from man-power project",
                "Output:", "  Original CSV with extra information for using Pivot Table",
            },
            // "J.Create Report, Update Report and TC Summary (A+H)",
            new String[] 
            {
                "Create Report and update them, also update TC Linked Issue ==> Report (A+H)", 
                "Input:",  "  Jira Bug & TC file, Template (for Test case output), and Input Excel File",
                "Output:", "  Updated reports specified within Input Excel File and TC summary with Linked issues",
            },
            //  "K.Update Report and TC summary (linked issue)",
            new String[] 
            {
                "Update report and TC Linked Issue (linked issue only)", 
                "Input:",  "  Jira Bug & TC file, Template (for Test case output), and root-directory of reports to be updated",
                "Output:", "  Updated reports under directories (named by root-directory-plus-datetime) and TC summary with Linked issues",
            },
            // "out-of-boundary text",
            new String[] 
            {
                "New feature to generate a better report (to-be-finalized)", 
                "Input:",  "  report input or template input (to-be-finalized)",
                "Output:", "  enhanced report (to-be-finalized)",
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
                "Test Report Path",
                "Main Test Plan",
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
            // "A.Create CST Report",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Input Excel File",
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
            // "D.Copy Report Only",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Input Excel File",
            },
            // "E.Remove Internal Sheets from Report",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Input Excel File",
            },
            // "F.Update Test Group Summary Report",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "G.Update Report Linked Issue",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "H.Update Keyword Rerpot and TC summary (7+9)",  
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "I.Man-Power Processing",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Exported CSV File",
            },
            // "J.Create Report, Update Report and TC Summary (A+H)",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Input Excel File",         // this is file selection
                "TC Template File",
            },
            // "K.Update Report and TC summary (linked issue)",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // out-of-boundary-text
            new String[] 
            {
                "Jira Bug File (OOB)", 
                "Jira TC File (OOB)",
                "Test Report Path (OOB)",
                "TC Template File (OOB)",
            },
        };

        public static String GetReportDescription(ReportType report_type)
        {
            String ret_str = "";
            ret_str += ReportGeneratorVersionString + "\r\n";
            int report_type_index = (int)report_type;
            foreach (String str in ReportDescription[report_type_index])
            {
                ret_str += str + "\r\n";
            }
            return ret_str;
        }

        public static String[] GetLabelTextArray(ReportType report_type)
        {
            int report_type_index = (int)report_type;
            return UI_Label[report_type_index];
        }

        private static List<String> SplitCommaSeparatedStringIntoList(String input_string)
        {
            List<String> ret_list = new List<String>();
            String[] csv_separators = { "," };
            if (String.IsNullOrWhiteSpace(input_string) == false)
            {
                // Separate keys into string[]
                String[] issues = input_string.Split(csv_separators, StringSplitOptions.RemoveEmptyEntries);
                if (issues != null)
                {
                    // string[] to List<String> (trimmed) and return
                    foreach (String str in issues)
                    {
                        ret_list.Add(str.Trim());
                    }
                }
            }
            return ret_list;
        }

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
            Issue.SheetName = XMLConfig.ReadAppSetting_String("BugList_ExportedSheetName");
            Issue.NameDefinitionRow = XMLConfig.ReadAppSetting_int("Issue_Row_NameDefine");
            Issue.DataBeginRow = XMLConfig.ReadAppSetting_int("Issue_Row_DataBegin");

            // config for test-case
            TestCase.KeyPrefix = XMLConfig.ReadAppSetting_String("TC_Key_Prefix");
            TestCase.SheetName = XMLConfig.ReadAppSetting_String("TCList_ExportedSheetName");
            TestCase.TemplateSheetName = XMLConfig.ReadAppSetting_String("TC_SheetName");
            TestCase.NameDefinitionRow = XMLConfig.ReadAppSetting_int("TC_Row_NameDefine");
            TestCase.DataBeginRow = XMLConfig.ReadAppSetting_int("TC_Row_DataBegin");

            // Status string to decompose into list of string
            // Begin
            //List<String> ret_list = new List<String>();
            //String links = XMLConfig.ReadAppSetting_String("LinkIssueFilterStatusString");
            //if ((links != null) && (links != ""))
            //{
            //    // Separate keys into string[]
            //    String[] issues = links.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            //    if (issues != null)
            //    {
            //        // string[] to List<String> (trimmed) and return
            //        foreach (String str in issues)
            //        {
            //            ret_list.Add(str.Trim());
            //        }
            //    }
            //}
            //ReportGenerator.fileter_status_list = ret_list;
            String links = XMLConfig.ReadAppSetting_String("LinkIssueFilterStatusString");
            ReportGenerator.filter_status_list_linked_issue = SplitCommaSeparatedStringIntoList(links);
            links = XMLConfig.ReadAppSetting_String("TestReport_FilterStatusString");
            KeywordReport.filter_status_list = SplitCommaSeparatedStringIntoList(links);
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

            // config for excel report output (also linked issue)
            StyleString.default_font = XMLConfig.ReadAppSetting_String("default_report_Font");
            StyleString.default_size = XMLConfig.ReadAppSetting_int("default_report_FontSize");
            StyleString.default_color = XMLConfig.ReadAppSetting_Color("default_report_FontColor");
            StyleString.default_fontstyle = XMLConfig.ReadAppSetting_FontStyle("default_report_FontStyle");
            // end config for excel report output

            // linked issue color
            ReportGenerator.LinkIssue_report_Font = XMLConfig.ReadAppSetting_String("LinkIssue_report_Font");
            ReportGenerator.LinkIssue_report_FontSize = XMLConfig.ReadAppSetting_int("LinkIssue_report_FontSize");
            ReportGenerator.LinkIssue_report_FontColor = XMLConfig.ReadAppSetting_Color("LinkIssue_report_FontColor");
            ReportGenerator.LinkIssue_report_FontStyle = XMLConfig.ReadAppSetting_FontStyle("LinkIssue_report_FontStyle");
            ReportGenerator.LinkIssue_A_Issue_Color = XMLConfig.ReadAppSetting_Color("LinkIssue_A_Issue_Color");
            ReportGenerator.LinkIssue_B_Issue_Color = XMLConfig.ReadAppSetting_Color("LinkIssue_B_Issue_Color");
            ReportGenerator.LinkIssue_C_Issue_Color = XMLConfig.ReadAppSetting_Color("LinkIssue_C_Issue_Color");
            ReportGenerator.LinkIssue_D_Issue_Color = XMLConfig.ReadAppSetting_Color("LinkIssue_D_Issue_Color");
            ReportGenerator.LinkIssue_WAIVED_ISSUE_COLOR = ReportGenerator.LinkIssue_report_FontColor;
            ReportGenerator.LinkIssue_CLOSED_ISSUE_COLOR = ReportGenerator.LinkIssue_report_FontColor;

            // Input Excel
            HeaderTemplate.SheetName_HeaderTemplate = XMLConfig.ReadAppSetting_String("InputExcel_Sheetname_HeaderTemplate");
            HeaderTemplate.SheetName_ReportList = XMLConfig.ReadAppSetting_String("InputExcel_Sheetname_ReportList");

            // config for keyword report
            Issue.Keyword_report_Font = XMLConfig.ReadAppSetting_String("Keyword_report_Font");
            Issue.Keyword_report_FontSize = XMLConfig.ReadAppSetting_int("Keyword_report_FontSize");
            Issue.Keyword_report_FontColor = XMLConfig.ReadAppSetting_Color("Keyword_report_FontColor");
            Issue.Keyword_report_FontStyle = XMLConfig.ReadAppSetting_FontStyle("Keyword_report_FontStyle");
            Issue.Keyword_A_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_A_Issue_Color");
            Issue.Keyword_B_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_B_Issue_Color");
            Issue.Keyword_C_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_C_Issue_Color");
            Issue.Keyword_D_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("Keyword_report_D_Issue_Color");
            Issue.Keyword_WAIVED_ISSUE_COLOR = Issue.Keyword_report_FontColor;
            Issue.Keyword_CLOSED_ISSUE_COLOR = Issue.Keyword_report_FontColor;

            KeywordReport.Replace_Conclusion = XMLConfig.ReadAppSetting_Boolean("Keyword_report_replace_conclusion");
            KeywordReport.Hide_Keyword_Result_Bug = XMLConfig.ReadAppSetting_Boolean("Keyword_report_Hide_Keyword_Result_Bug");
            KeywordReport.Auto_Correct_Sheetname = XMLConfig.ReadAppSetting_Boolean("Keyword_Auto_Correct_Worksheet");

            // config for report C
            KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = XMLConfig.ReadAppSetting_Boolean("Report_C_CopyFileOnly");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal = XMLConfig.ReadAppSetting_Boolean("Report_C_Remove_AUO_Internal");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = XMLConfig.ReadAppSetting_Boolean("Report_C_Update_Full_Header");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template = XMLConfig.ReadAppSetting_Boolean("Report_C_Update_Header_by_Template");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion = XMLConfig.ReadAppSetting_Boolean("Report_C_Replace_Conclusion");

            KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = XMLConfig.ReadAppSetting_Boolean("Report_C_Update_Report_Sheetname");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = XMLConfig.ReadAppSetting_Boolean("Report_C_Clear_Keyword_Result");
            KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row = XMLConfig.ReadAppSetting_Boolean("Report_C_Hide_Keyword_Result_Bug_Row");
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

        private static ReportType default_selected_report_type = ReportSelectableTable[0];
        //private static ReportType default_selected_report_type = ReportType.Man_Power_Processing;

        private void InitializeReportFunctionListBox()
        {
            foreach (ReportType report_type in ReportSelectableTable)
            {
                String report_name;
                report_name = GetReportName(report_type);
                comboBoxReportSelect.Items.Add(report_name);
            }
            //int default_select_index = (int)ReportType.FullIssueDescription_Summary; // current default
            Set_comboBoxReportSelect_ByReportType(default_selected_report_type);
            UpdateUIAfterReportTypeChanged(default_selected_report_type);
        }

        //private Boolean Set_comboBoxReportSelect_IndexValue(int index_value)
        //{
        //    Boolean b_ret = false;
        //    if ((index_value >= 0) && (index_value < ReportTypeCount))
        //    {
        //        comboBoxReportSelect.SelectedIndex = index_value;
        //        b_ret = true;
        //    }
        //    return b_ret;
        //}

        //private int Get_comboBoxReportSelect_IndexValue()
        //{
        //    int selected_index = comboBoxReportSelect.SelectedIndex;
        //    if ((selected_index < 0) || (selected_index >= ReportTypeCount))
        //    {
        //        // shouldn't be out of range.
        //        return 0;
        //    }
        //    else
        //    {
        //        return comboBoxReportSelect.SelectedIndex;
        //    }
        //}

        private void Set_comboBoxReportSelect_ByReportType(ReportType report_type)
        {
            for (int index = 0; index < ReportTypeCount; index++)
            {
                if (ReportSelectableTable[index] == report_type)
                {
                    comboBoxReportSelect.SelectedIndex = index;
                    //MsgWindow.AppendText(@"Report Type " + '"' + report_type.ToString() + '"' + " is selected.\n");
                    MsgWindow.AppendText(@"Report Type " + '"' + GetReportName(report_type) + '"' + " is selected.\n");
                    return;
                }
            }
            // shouldn't be here, need to debug
            MsgWindow.AppendText("WARNING: Set_comboBoxReportSelect_ByReportType Report Type not found.\n");
            comboBoxReportSelect.SelectedIndex = 0;
        }
        private ReportType Get_comboBoxReportSelect_ReturnReportType()
        {
            // comboBoxReportSelect.SelectedIndex range check should be done when it is set.
            int selected_index = comboBoxReportSelect.SelectedIndex;
            if ((selected_index < 0) || (selected_index >= ReportTypeCount))
            {
                // shouldn't be out of range.
                // to-be-checked here.
                MsgWindow.AppendText("WARNING: Get_comboBoxReportSelect_ReturnReportType out-of-range error.\n");
                comboBoxReportSelect.SelectedIndex = 0;
            }
            return ReportSelectableTable[comboBoxReportSelect.SelectedIndex];
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
            if (!string.IsNullOrEmpty(strCompTime))
            {
                string[] subs = strCompTime.Split(' ', '/'); // use ' ' & '/' as separator
                strBuildDate = "(" + subs[0] + subs[1] + subs[2] + ")";
            }
            ReportGeneratorVersionString = "ReportGenerator_V" + version + strBuildDate;
            this.Text = ReportGeneratorVersionString;

            LoadConfigAll();

            InitializeReportFunctionListBox();

            if ((Storage.GetFullPath(txtBugFile.Text) == "") ||
                (Storage.GetFullPath(txtTCFile.Text) == "") ||
                (Storage.GetFullPath(txtReportFile.Text) == "") ||
                (Storage.GetFullPath(txtOutputTemplate.Text) == ""))
            {
                MsgWindow.AppendText("WARNING: one or more sample files do not exist.\n");
            }
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
                ReportGenerator.UpdateGlobalIssueList(Issue.GenerateIssueList(buglist_filename));
                //ReportGenerator.lookup_BugList = Issue.UpdateIssueListLUT(ReportGenerator.global_issue_list);
                // update LUT
                MsgWindow.AppendText("bug_list finished!\n");
                return true;
            }
        }

        private Boolean ReadGlobalTCListTask(String filename)
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
                List<TestCase> new_tc_list = TestCase.GenerateTestCaseList_New(tclist_filename);
                ReportGenerator.UpdateGlobalTestcaseList(new_tc_list);
                MsgWindow.AppendText("tc_list finished!\n");
                return true;
            }
        }

        private bool LoadIssueListIfEmpty(String filename)
        {
            if (ReportGenerator.IsGlobalIssueListEmpty())
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
            ReportGenerator.ClearGlobalIssueList();
            //ReportGenerator.lookup_BugList.Clear();
            KeywordReport.ClearGlobalKeywordList();
        }

        private Boolean LoadTCListIfEmpty(String filename)
        {
            if (ReportGenerator.IsGlobalTestcaseListEmpty())
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
            ReportGenerator.ClearGlobalTestcaseList();
            ReportGenerator.ClearTestcaseLUT_by_Key();
            ReportGenerator.ClearTestcaseLUT_by_Sheetname();
            KeywordReport.ClearGlobalKeywordList();
        }

        private bool Execute_WriteIssueDescriptionToTC(String tc_file, String template_file, String buglist_file, String judgement_report_dir = "")
        {
            if ((ReportGenerator.IsGlobalIssueListEmpty()) || (ReportGenerator.IsGlobalTestcaseListEmpty()) ||
                (!Storage.FileExists(tc_file)) || (!Storage.FileExists(template_file) || (!Storage.FileExists(buglist_file)))
                || ((judgement_report_dir != "") && !Storage.DirectoryExists(judgement_report_dir)))
            {
                // protection check
                // Bug/TC files must have been loaded
                return false;
            }

            // This full issue description is needed for report purpose
            //Dictionary<string, List<StyleString>> global_issue_description_list = StyleString.GenerateIssueDescription(ReportGenerator.global_issue_list);
            //Dictionary<string, List<StyleString>> global_issue_description_list_severity =
            //            StyleString.GenerateIssueDescription_Severity_by_Linked_Issue(ReportGenerator.ReadGlobalIssueList());
            //List<TestCase> before = ReportGenerator.ReadGlobalTestcaseList();
            //List<TestCase> after = TestCase.UpdateTCLinkedIssueList(before, ReportGenerator.ReadGlobalIssueList(), global_issue_description_list_severity);
            //ReportGenerator.UpdateGlobalTestcaseList(after);

            //            ReportGenerator.WriteBacktoTCJiraExcel(tc_file);
            //ReportGenerator.WriteBacktoTCJiraExcelV2(tc_file, template_file, judgement_report_dir);
            //ReportGenerator.WriteBacktoTCJiraExcelV3(tc_file, template_file, bug_file, ReportGenerator.ReadGlobalIssueList(), global_issue_description_list_severity, judgement_report_dir);
            ReportGenerator.WriteBacktoTCJiraExcelV3(tclist_filename: tc_file, template_filename: template_file, buglist_file: buglist_file, judgement_report_dir: judgement_report_dir);
            return true;
        }

        private bool Execute_WriteIssueDescriptionToSummary(String template_file)
        {
            if ((ReportGenerator.IsGlobalIssueListEmpty()) || (ReportGenerator.IsGlobalTestcaseListEmpty()) ||
                (!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

            // This full issue description is needed for report purpose
            //Dictionary<string, List<StyleString>> global_full_issue_description_list =
            //                            StyleString.GenerateFullIssueDescription(ReportGenerator.ReadGlobalIssueList());

            //SummaryReport.SaveIssueToSummaryReport(template_file, global_full_issue_description_list);
            SummaryReport.SaveIssueToSummaryReport(template_file);

            return true;
        }

        //private bool Execute_CreateStandardTestReportTask(String main_file)
        //{
        //    if (!Storage.FileExists(main_file))
        //    {
        //        // protection check
        //        return false;
        //    }

        //    return TestReport.CreateStandardTestReportTask(main_file);
        //}

        //private bool Execute_CreateTestReportbyTestCaseTask(String report_src_dir, String output_report_dir)
        //{
        //    if (!Storage.DirectoryExists(report_src_dir) || !Storage.DirectoryExists(output_report_dir))
        //    {
        //        // protection check
        //        // source_dir & output_dir must exist.
        //        return false;
        //    }

        //    return TestReport.CopyTestReportbyTestCase(report_src_dir, output_report_dir);
        //}

        private bool Execute_KeywordIssueGenerationTask(String FileOrDirectoryName, Boolean IsDirectory = false)
        {
            String output_report_path;  // not used for this task
            return Execute_KeywordIssueGenerationTask_returning_report_path(FileOrDirectoryName, IsDirectory, out output_report_path);
        }

        private bool Execute_KeywordIssueGenerationTask_returning_report_path(String FileOrDirectoryName, Boolean IsDirectory, out String output_report_path)
        {
            List<String> file_list = new List<String>();
            String source_dir;
            output_report_path = "";
            if (IsDirectory == false)
            {
                if ((ReportGenerator.IsGlobalIssueListEmpty()) || (!Storage.FileExists(FileOrDirectoryName)))
                {
                    // protection check
                    return false;
                }
                file_list.Add(FileOrDirectoryName);
                source_dir = Storage.GetDirectoryName(FileOrDirectoryName);
            }
            else
            {
                if ((ReportGenerator.IsGlobalIssueListEmpty()) || (!Storage.DirectoryExists(FileOrDirectoryName)))
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

            //// this is for keyword report, how to input linked issue report list???
            //Dictionary<string, List<StyleString>> global_issue_description_list_severity =
            //                    StyleString.GenerateIssueDescription_Keyword_Issue(ReportGenerator.ReadGlobalIssueList());
            String out_dir = KeywordReport.TestReport_Default_Output_Dir;
            if ((out_dir != "") && Storage.DirectoryExists(out_dir))
            {
                output_report_path = KeywordReport.TestReport_Default_Output_Dir;
            }
            else
            {
                output_report_path = Storage.GenerateDirectoryNameWithDateTime(source_dir);
            }
            KeywordReport.KeywordIssueGenerationTaskV4(report_list, source_dir, output_report_path);
            return true;
        }

        private bool Execute_KeywordIssueGenerationTask_returning_report_path_simplified(String FileOrDirectoryName, Boolean IsDirectory, out String output_report_path)
        {
            List<String> file_list = new List<String>();
            String source_dir;
            output_report_path = "";
            if (IsDirectory == false)
            {
                if ((ReportGenerator.IsGlobalIssueListEmpty()) || (!Storage.FileExists(FileOrDirectoryName)))
                {
                    // protection check
                    return false;
                }
                file_list.Add(FileOrDirectoryName);
                source_dir = Storage.GetDirectoryName(FileOrDirectoryName);
            }
            else
            {
                if ((ReportGenerator.IsGlobalIssueListEmpty()) || (!Storage.DirectoryExists(FileOrDirectoryName)))
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

            //// this is for keyword report, how to input linked issue report list???
            //Dictionary<string, List<StyleString>> global_issue_description_list_severity =
            //                    StyleString.GenerateIssueDescription_Keyword_Issue(ReportGenerator.ReadGlobalIssueList());
            String out_dir = KeywordReport.TestReport_Default_Output_Dir;
            if ((out_dir != "") && Storage.DirectoryExists(out_dir))
            {
                output_report_path = KeywordReport.TestReport_Default_Output_Dir;
            }
            else
            {
                output_report_path = Storage.GenerateDirectoryNameWithDateTime(source_dir);
            }
            KeywordReport.KeywordIssueGenerationTaskV4_simplified(report_list, source_dir, output_report_path);
            return true;
        }

        //private bool Execute_KeywordIssueGenerationTask_returning_report_path_update_bug_list(String FileOrDirectoryName, Boolean IsDirectory,  out String output_report_path)
        //{
        //    List<String> file_list = new List<String>();
        //    String source_dir;
        //    output_report_path = "";
        //    if (IsDirectory == false)
        //    {
        //        if ((ReportGenerator.IsGlobalIssueListEmpty()) || (!Storage.FileExists(FileOrDirectoryName)))
        //        {
        //            // protection check
        //            return false;
        //        }
        //        file_list.Add(FileOrDirectoryName);
        //        source_dir = Storage.GetDirectoryName(FileOrDirectoryName);
        //    }
        //    else
        //    {
        //        if ((ReportGenerator.IsGlobalIssueListEmpty()) || (!Storage.DirectoryExists(FileOrDirectoryName)))
        //        {
        //            // protection check
        //            return false;
        //        }
        //        file_list = Storage.ListFilesUnderDirectory(FileOrDirectoryName);
        //        //MsgWindow.AppendText("File found under directory " + FileOrDirectoryName + "\n");
        //        //foreach (String filename in file_list)
        //        //    MsgWindow.AppendText(filename + "\n");
        //        source_dir = FileOrDirectoryName;
        //    }
        //    // filename check to exclude non-report files.
        //    //List<String> report_list = Storage.FilterFilename(file_list);
        //    // NOTE: FilterFilename() execution is now relocated within KeywordIssueGenerationTaskV4()
        //    List<String> report_list = file_list;

        //    // This issue description is needed for report purpose
        //    //ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);
        //    Dictionary<string, List<StyleString>> global_issue_description_list_severity =
        //                        StyleString.GenerateIssueDescription_Keyword_Issue(ReportGenerator.ReadGlobalIssueList());
        //    String out_dir = KeywordReport.TestReport_Default_Output_Dir;
        //    if ((out_dir != "") && Storage.DirectoryExists(out_dir))
        //    {
        //        output_report_path = KeywordReport.TestReport_Default_Output_Dir;
        //    }
        //    else
        //    {
        //        output_report_path = Storage.GenerateDirectoryNameWithDateTime(source_dir);
        //    }
        //    KeywordReport.KeywordIssueGenerationTaskV4(report_list, global_issue_description_list_severity, source_dir, output_report_path);
        //    return true;
        //}

        private bool Execute_FindFailTCLinkedIssueAllClosed(String tc_file, String template_file)
        {
            if ((ReportGenerator.IsGlobalIssueListEmpty()) || (ReportGenerator.IsGlobalTestcaseListEmpty()) ||
                (!Storage.FileExists(tc_file)) || (!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

            //// This issue description is needed for report purpose
            //Dictionary<string, List<StyleString>> global_issue_description_list = StyleString.GenerateIssueDescription(ReportGenerator.ReadGlobalIssueList());

            ReportGenerator.FindFailTCLinkedIssueAllClosed(tc_file, template_file, ReportGenerator.ReadGlobalIssueList());
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

            TestReport.AutoCorrectReport_by_Folder(report_root: report_root, Output_dir: Storage.GenerateDirectoryNameWithDateTime(report_root));

            return true;
        }

        private bool Execute_AutoCorrectTestReportByExcel_Task(String excel_input_file)
        {
            if (!Storage.FileExists(excel_input_file))
            {
                // protection check
                return false;
            }

            CopyReport.UpdateTestReportByOptionAndSaveAsAnother(excel_input_file);

            return true;
        }

        private bool Execute_AutoCorrectTestReportByExcel_with_output_report_list_Task(String excel_input_file, out List<String> output_report_list)
        {
            output_report_list = new List<String>();

            if (!Storage.FileExists(excel_input_file))
            {
                // protection check
                return false;
            }

            CopyReport.UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(excel_input_file, out output_report_list);

            return true;
        }

        private bool Execute_Man_Power_Processing_Task(String excel_input_file)
        {
            if (!Storage.FileExists(excel_input_file))
            {
                // protection check
                return false;
            }

            ManPowerTask.ProcessManPowerPlan_V2(excel_input_file);

            return true;
        }

        private bool Execute_UpdaetGroupSummaryReport_Task(String report_path)
        {
            if (!Storage.DirectoryExists(report_path))
            {
                // protection check
                return false;
            }

            KeywordReport.Update_Group_Summary(report_path);

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
            ReportType report_type = Get_comboBoxReportSelect_ReturnReportType();
            bool sel_file = false;
            String init_dir;
            switch (report_type)
            {
                case ReportType.KeywordIssue_Report_SingleFile:
                case ReportType.Update_Repoart_A_then_Report_H:              //Report J = A + H
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
            ReportType report_type = Get_comboBoxReportSelect_ReturnReportType();
            bool sel_file = true;
            String init_dir;
            //switch (report_type)
            //{
            //    case ReportType.CreateCSTReport:
            //        //case ReportType.FindAllKeywordInReport:
            //        sel_file = false;  // Here select directory instead of file
            //        init_dir = txtOutputTemplate.Text;
            //        break;
            //    default:
            //        // default is file selection here.
            //        init_dir = Storage.GetFullPath(txtOutputTemplate.Text);
            //        break;
            //}
            init_dir = Storage.GetFullPath(txtOutputTemplate.Text);

            if (report_type == ReportType.Man_Power_Processing)
            {
                String ret_str = Storage.UsesrSelectCSVFilename(init_dir);
                if (ret_str != "")
                {
                    txtOutputTemplate.Text = ret_str;
                    btnSelectOutputTemplate_Clicked = true;
                }
            }
            else
            {
                String ret_str = SelectDirectoryOrFile(init_dir, sel_file);
                if (ret_str != "")
                {
                    txtOutputTemplate.Text = ret_str;
                    btnSelectOutputTemplate_Clicked = true;
                }
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
        private Boolean UpdateTextBoxPathToFullAndCheckExist(ref TextBox box)
        {
            String name = Storage.GetFullPath(box.Text);
            if (!Storage.FileExists(name))
            {
                MsgWindow.AppendText("Files can't be found and please check again: " + box.Text + "\n");
                return false;
            }
            box.Text = name;
            return true;
        }

        private Boolean UpdateTextBoxDirToFullAndCheckExist(ref TextBox box)
        {
            String name = Storage.GetFullPath(box.Text);
            if (!Storage.DirectoryExists(name))
            {
                MsgWindow.AppendText("Directory can't be found and please check again: " + box.Text + "\n");
                return false;
            }
            box.Text = name;
            return true;
        }

        static public Queue<String> SYSTEM_LOG = new Queue<String>();
        static public void SystemLogAdd(String log_string) { SYSTEM_LOG.Enqueue(log_string); }
        static public void SystemLogAddLine(String log_string) { SystemLogAdd(log_string + "\r\n"); }

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            bool bRet;

            ReportType report_type = Get_comboBoxReportSelect_ReturnReportType();
            ClearIssueList();
            ClearTCList();

            UpdateUIDuringExecution(report_type: report_type, executing: true);

            MsgWindow.AppendText("Executing: " + GetReportName(report_type) + ".\n");

            Boolean open_excel_ok = ExcelAction.OpenExcelApp();
            if (open_excel_ok)
            {
                Stack<Boolean> temp_option_stack = new Stack<Boolean>();

                // Must be updated if new report type added #NewReportType
                switch (report_type)
                {
                    case ReportType.FullIssueDescription_TC:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        //if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;  // can be removed after v2
                        //if (!LoadTCListIfEmpty(txtTCFile.Text)) break;  // can be removed after v2
                        bRet = ReportGenerator.Execute_ExtendLinkIssueAndUpdateStatusByLinkIssueFilteredCount_v2(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text);
                        break;
                    case ReportType.FullIssueDescription_Summary: // report 2 not used now
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_WriteIssueDescriptionToSummary(template_file: txtOutputTemplate.Text);
                        break;
                    case ReportType.CreateImportToJiraCSV:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        //bRet = Execute_CreateStandardTestReportTask(main_file: txtOutputTemplate.Text);
                        break;
                    case ReportType.KeywordIssue_Report_SingleFile:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);    // File path here
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        bRet = Execute_KeywordIssueGenerationTask(FileOrDirectoryName: txtReportFile.Text, IsDirectory: true);
                        break;
                    case ReportType.KeywordIssue_Report_Directory:                  // REport 7
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);     // Directory path here
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        bRet = Execute_KeywordIssueGenerationTask(FileOrDirectoryName: txtReportFile.Text, IsDirectory: true);
                        break;
                    case ReportType.TC_Likely_Passed:           /// report 5 not used now
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_FindFailTCLinkedIssueAllClosed(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text);
                        break;
                    case ReportType.FindAllKeywordInReport:
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtStandardTestReport);
                        //String main_file = txtStandardTestReport.Text;
                        //String file_dir = Storage.GetDirectoryName(main_file);
                        String output_filename = "";//use default in config file
                        String report_root_dir = Storage.GetFullPath(txtReportFile.Text);
                        bRet = Execute_ListAllDetailedTestPlanKeywordTask(report_root: report_root_dir, output_file: output_filename);
                        break;
                    case ReportType.Excel_Sheet_Name_Update_Tool:
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);     // Directory path here
                        // bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: true);
                        bRet = true;
                        break;
                    case ReportType.FullIssueDescription_TC_report_judgement:           // Report 9
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_WriteIssueDescriptionToTC(tc_file: txtTCFile.Text, judgement_report_dir: txtReportFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text);
                        break;
                    case ReportType.CreateCSTReport:                                    // Report A
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        //// based on tc, create report structure
                        //if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        ////String dest_dir = Storage.GetFullPath(txtReportFile.Text),
                        ////       src_dir = Storage.GetFullPath(txtOutputTemplate.Text);
                        //String src_dir = Storage.GetFullPath(txtReportFile.Text),
                        //       dest_dir = Storage.GetFullPath(txtOutputTemplate.Text);
                        //bRet = Execute_CreateTestReportbyTestCaseTask(report_src_dir: src_dir, output_report_dir: dest_dir);
                        // update report according to jira bug
                        //if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        // to-be-implemented

                        // update judgement and header
                        // to-be-implemented

                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        // to-be-updated
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template);
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = true;
                        //KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row = false;
                        //KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template = true;
                        bRet = Execute_AutoCorrectTestReportByExcel_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = temp_option_stack.Pop();
                        break;
                    case ReportType.TC_AutoCorrectReport_By_Filename:
                        UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile);
                        bRet = Execute_AutoCorrectTestReportByFilename_Task(report_root: Storage.GetFullPath(txtReportFile.Text));
                        break;
                    case ReportType.TC_AutoCorrectReport_By_ExcelList:                          // Report C
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        // to-be-updated
                        bRet = Execute_AutoCorrectTestReportByExcel_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        break;
                    case ReportType.CopyReportOnly:                                             // Report D
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        // copy files only
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly);
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = true;
                        bRet = Execute_AutoCorrectTestReportByExcel_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = temp_option_stack.Pop();
                        break;
                    case ReportType.RemoveInternalSheet:                                        // Report E
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        // copy files only
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template);
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal = true;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template = false;
                        bRet = Execute_AutoCorrectTestReportByExcel_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = temp_option_stack.Pop();
                        break;
                    case ReportType.TC_GroupSummaryReport:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        // update tc_linked_issue_description
                        //Dictionary<string, List<StyleString>> TC_issue_description =
                        //            StyleString.GenerateIssueDescription_Severity_by_Linked_Issue(ReportGenerator.ReadGlobalIssueList());
                        //List<TestCase> before = ReportGenerator.ReadGlobalTestcaseList();
                        //List<TestCase> after = TestCase.UpdateTCLinkedIssueList(before, ReportGenerator.ReadGlobalIssueList(), TC_issue_description);
                        //ReportGenerator.UpdateGlobalTestcaseList(after);

                        bRet = Execute_UpdaetGroupSummaryReport_Task(report_path: txtReportFile.Text);
                        break;
                    case ReportType.Update_Report_Linked_Issue:
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        //bRet = Execute_CreateTestReportbyTestCaseTask(report_src_dir: src_dir, output_report_dir: dest_dir);
                        break;
                    case ReportType.Update_Keyword_and_TC_Report: // Report H
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;

                        String report_output_path;
                        bRet = Execute_KeywordIssueGenerationTask_returning_report_path(txtReportFile.Text, true, out report_output_path);
                        bRet = Execute_WriteIssueDescriptionToTC(tc_file: txtTCFile.Text, judgement_report_dir: report_output_path, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text);
                        break;
                    case ReportType.Man_Power_Processing:
                        tabInfomation.SelectedTab = tabInfomation.TabPages[1];
                        Application.DoEvents();
                        //Thread.Sleep(500);
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate))
                        {
                            MsgWindow.AppendText("Using the following file: " + txtOutputTemplate.Text + "\n");
                            bRet = Execute_Man_Power_Processing_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                            if (!bRet)
                            {
                                MsgWindow.AppendText("Execution was not unsuccessful and please check CSV file\n");
                            }
                        }
                        break;
                    case ReportType.Update_Repoart_A_then_Report_H:                                    //Report J = A + H
                        //
                        //
                        // TO BE UPDATED
                        //
                        //
                        // copied from report A
                        // NOTE: Input Excel File is storted in txtReportFile for Report J
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        String report_j_input_excel_selected = txtOutputTemplate.Text;
                        // to-be-updated
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header);
                        temp_option_stack.Push(KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template);
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = true;
                        //KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row = false;
                        //KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = false;
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template = true;
                        List<String> report_list;
                        bRet = Execute_AutoCorrectTestReportByExcel_with_output_report_list_Task(Storage.GetFullPath(report_j_input_excel_selected), out report_list);
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal = temp_option_stack.Pop();
                        KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly = temp_option_stack.Pop();

                        // copied from report H
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        //UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;

                        String report_output_path_2;
                        bRet = Execute_KeywordIssueGenerationTask_returning_report_path(txtReportFile.Text, true, out report_output_path_2);
                        bRet = Execute_WriteIssueDescriptionToTC(tc_file: txtTCFile.Text, judgement_report_dir: report_output_path_2, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text);


                        break;
                    case ReportType.Update_Report_Linked_Issue_and_TC_Report: // Report K
                        UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile);
                        UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate);
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;

                        String report_output_path_report_k;
                        bRet = Execute_KeywordIssueGenerationTask_returning_report_path_simplified(txtReportFile.Text, true, out report_output_path_report_k);
                        bRet = Execute_WriteIssueDescriptionToTC(tc_file: txtTCFile.Text, judgement_report_dir: report_output_path_report_k, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text);
                        break;
                    default:
                        MsgWindow.AppendText("Report Type has exception. Please check\n");
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
            UpdateUIDuringExecution(report_type: report_type, executing: false);
        }

        private void SetEnable_BugFile(bool value)
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
            ReportType selected_report_type = Get_comboBoxReportSelect_ReturnReportType();
            UpdateUIAfterReportTypeChanged(selected_report_type);
            String[] string_array = GetLabelTextArray(selected_report_type);
            label_issue.Text = string_array[0];
            label_tc.Text = string_array[1];
            label_1st.Text = string_array[2];
            label_2nd.Text = string_array[3];
            tabInfomation.SelectedTab = tabInfomation.TabPages[0];
        }

        private void UpdateUIDuringExecution(ReportType report_type, bool executing)
        {
            if (!executing)
            {
                UpdateFilenameBoxUIForReportType(report_type);
                btnCreateReport.Enabled = true;
            }
            else
            {
                SetEnable_BugFile(false);
                SetEnable_TCFile(false);
                SetEnable_ReportFile(false);
                SetEnable_OutputTemplate(false);
                btnCreateReport.Enabled = false;
            }
        }

        private void UpdateFilenameBoxUIForReportType(ReportType report_type)
        {
            // Must be updated if new report type added #NewReportType
            switch (report_type)
            {
                case ReportType.FullIssueDescription_TC: // "1.Issue Description for TC"
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.FullIssueDescription_Summary: // "2.Issue Description for Summary"
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.CreateImportToJiraCSV:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.KeywordIssue_Report_SingleFile:
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.KeywordIssue_Report_Directory:
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.TC_Likely_Passed:
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.FindAllKeywordInReport:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.Excel_Sheet_Name_Update_Tool:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.FullIssueDescription_TC_report_judgement: // "1.Issue Description for TC"
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.CreateCSTReport:            // Report A
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.TC_AutoCorrectReport_By_Filename:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.TC_AutoCorrectReport_By_ExcelList:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.CopyReportOnly:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.RemoveInternalSheet:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.TC_GroupSummaryReport:
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(false);
                    break;
                case ReportType.Update_Report_Linked_Issue:
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.Update_Keyword_and_TC_Report:
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.Man_Power_Processing:
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.Update_Repoart_A_then_Report_H:            // Report J = A + H
                    // copied from report H -- all enabled
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.Update_Report_Linked_Issue_and_TC_Report:   // Report K
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                default:
                    // Shouldn't be here
                    break;
            }
        }

        private void UpdateUIAfterReportTypeChanged(ReportType report_type)
        {
            txtReportInfo.Text = GetReportDescription(report_type);
            UpdateFilenameBoxUIForReportType(report_type);


            // Must be updated if new report type added #NewReportType
            switch (report_type)
            {
                case ReportType.FullIssueDescription_TC: // "1.Issue Description for TC"
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.FullIssueDescription_Summary: // "2.Issue Description for Summary"
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_Summary");
                    break;
                case ReportType.CreateImportToJiraCSV:
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
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
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
                case ReportType.CreateCSTReport:                  // report A -- tailored version of report C
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("Report_A_Default_Excel");
                    break;
                case ReportType.TC_AutoCorrectReport_By_Filename:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    break;
                case ReportType.TC_AutoCorrectReport_By_ExcelList:      // Report C
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("Report_C_Default_Excel");
                    break;
                case ReportType.CopyReportOnly:                         // Report D
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("Report_D_Copy_Only_Default_Excel");
                    break;
                case ReportType.RemoveInternalSheet:                    // Report E
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("Report_E_Remove_AUO_Sheet_Default_Excel");
                    break;
                case ReportType.TC_GroupSummaryReport:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    break;
                case ReportType.Update_Report_Linked_Issue:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.Update_Keyword_and_TC_Report: // original adopted from report 9     -- report H
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.Man_Power_Processing:
                    if (!btnSelectOutputTemplate_Clicked)
                    {
                        String short_userName = Storage.GetWindowsLoginUserName();
                        txtOutputTemplate.Text = @"C:\Users\" + short_userName + @"\Downloads\Advance Roadmaps.csv";
                    }
                    break;
                case ReportType.Update_Repoart_A_then_Report_H:                  // Report J = A + H
                    // copied from report H
                    // NOTE: Input Excel File is storted in txtReportFile for Report J
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Report_A_Default_Excel");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.Update_Report_Linked_Issue_and_TC_Report: // Report K
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Keyword_default_report_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
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
