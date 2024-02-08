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
            FullIssueDescription_TC = 0,                    // Report 1 -- Update TC linked issue and update Testcase Status according to linked issue (severity/count)
            FullIssueDescription_Summary,
            CreateImportToJiraCSV,
            KeywordIssue_Report_SingleFile,
            ReadReportContentVariable,                      // Report 5 -- read report variable and write to Excel
            WriteReportContentVariable,                     // Report 6 -- Read from Excel report variable and write to report
            KeywordIssue_Report_Directory,                  // Report 7
            Excel_Sheet_Name_Update_Tool,
            FullIssueDescription_TC_report_judgement,       // Report 9 -- Update TC linked issue, according to reports, update Test Purpose, Test Criteria and Testcase Status.
            CreateCSTReport,                                // Report A -- Create CST
            FinalCSTReport,                                 // Report B -- Create CST + update linked issue & judgement + copy worksheet at beginning/ending (options)
            ConfigurableReportUpdate,                       // Report C
            CopyReportOnly,                                 // Report D -- copy only version of report C
            RemoveInternalSheet,                            // Report E -- remove internalsheet version of report C
            TC_GroupSummaryReport,                          // Report F -- Not used
            Update_Report_Linked_Issue,
            Update_Keyword_and_TC_Report,                   // Report H -- it is report 7 + 9
            Man_Power_Processing,                           // Report I -- man-power
            Update_PassReport_and_TC_Summary,               // Report J -- 
            Update_Report_Linked_Issue_and_TC_Report,       // Report K -- simplified version of report H - no keyword at all, less format update/correction on report 7 + Report 9
            Update_Repoart_A_then_Report_K,                 // Report L -- it is report A + K
        }

        public static ReportType[] ReportSelectableTable =
        {
            ReportType.FullIssueDescription_TC,                     // report 1
            //ReportType.FullIssueDescription_Summary,
            //ReportType.StandardTestReportCreation,
            //ReportType.KeywordIssue_Report_SingleFile,                            // Hide since v 1.3.21.0
            ReportType.ReadReportContentVariable,                   // Report 5 -- read report variable and write to Excel
            ReportType.WriteReportContentVariable,                  // Report 6 -- Read from Excel report variable and write to report
            //ReportType.KeywordIssue_Report_Directory,               // report 7   // Hide since v 1.3.21.0
            //ReportType.Excel_Sheet_Name_Update_Tool,
            ReportType.FullIssueDescription_TC_report_judgement,    // report 9
            ReportType.CreateCSTReport,                             // report A
            ReportType.FinalCSTReport,                              // Report B -- Create CST + update linked issue & judgement + copy worksheet at beginning/ending (options)
            ReportType.ConfigurableReportUpdate, 
            ReportType.CopyReportOnly,
            ReportType.RemoveInternalSheet, 
            //ReportType.TC_GroupSummaryReport,
            //ReportType.Update_Report_Linked_Issue,
            //ReportType.Update_Keyword_and_TC_Report,                // Report H   // Hide since v 1.3.21.0
            //ReportType.Man_Power_Processing,
            ReportType.Update_PassReport_and_TC_Summary,              // Report J   // Hide since v 1.3.21.0
            ReportType.Update_Report_Linked_Issue_and_TC_Report,    // Report K
            ReportType.Update_Repoart_A_then_Report_K,              // Report L
        };

        //public static ReportType[] ReportSelectableTable =
        //{
        //    ReportType.FullIssueDescription_TC,
        //    ReportType.FullIssueDescription_Summary,
        //    ReportType.CreateImportToJiraCSV,
        //    ReportType.KeywordIssue_Report_SingleFile,
        //    ReportType.ReadReportContentVariable, 
        //    ReportType.WriteReportContentVariable,  
        //    ReportType.KeywordIssue_Report_Directory,
        //    ReportType.Excel_Sheet_Name_Update_Tool,
        //    ReportType.FullIssueDescription_TC_report_judgement,
        //    ReportType.CreateCSTReport,
        //    ReportType.FinalCSTReport,
        //    ReportType.ConfigurableReportUpdate,
        //    ReportType.CopyReportOnly,
        //    ReportType.RemoveInternalSheet, 
        //    ReportType.TC_GroupSummaryReport,
        //    ReportType.Update_Report_Linked_Issue,
        //    ReportType.Update_Keyword_and_TC_Report,
        //    ReportType.Man_Power_Processing,
        //    ReportType.Update_PassReport_and_TC_Summary,
        //    ReportType.Update_Report_Linked_Issue_and_TC_Report,
        //    ReportType.Update_Repoart_A_then_Report_K,
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
            "5.Read Report Header to Excel",
            "6.Write Report Header according to Excel",
            "7.Keyword Issue - Directory",
            "8.Excel sheet name update tool",
            "9.TC issue/judgement",
            "A.Create CST Report",
            "B.Create CST Report and update conclusion",
            "C.Configurable Report Processor",
            "D.Copy Report Only",
            "E.Remove Internal Sheets from Report",
            "F.Update Test Group Summary Report",
            "G.Update Report Linked Issue",
            "H.Update Keyword Report and TC summary (7+9)",
            "I.Man-Power Processing",
            "J.Update Passed Report and TC summary (linked issue)",
            "K.Update All Report and TC summary (linked issue)",
            "L.Create Report, Update Report and TC Summary (A+K)",
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
            // "5.Read Report Header to Excel",
            new String[] 
            {
                "Read report header based on header template file and output data to Excel", 
                "Input:",  "  Input excel file containing header template",
                "Output:", "  Excel worksheet which contains header data defined as header template variables",
            },
            // "6.Write Report Header according to Excel",
            new String[] 
            {
                "Write report header based on header template file and data on input Excel", 
                "Input:",  "  Input excel file containing header template and header data worksheet",
                "Output:", "  Header of reports are updated based on header template and header data",
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
            // "A.CST Test Report Creator",
            new String[] 
            {
                "Create CST Report", 
                "Input:",  "  Input excel file",
                "Output:", "  Reports copied and renamed (filename / worksheet name) according to input excel file",
            },
            // "B.CST Test Report Creator + update judgement link isssue...etc",
            new String[] 
            {
                "Create CST Report + linked issue / judgement...etc and update TC Linked Issue", 
                "Input:",  "  Jira Bug & TC file, Template (for Test case output), and Input Excel File",
                "Output:", "  Updated reports specified within Input Excel File and TC summary with Linked issues",
            },
            // "C.Create New Model Report",
            new String[] 
            {
                "Configurable Report Function", 
                "Input:",  "  Input excel file",
                "Output:", "  Reports copied and renamed (filename / worksheet name) and contents are updated according to input excel file and option/template",
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
                "Update report and TC Linked Issue (linked issue only)", 
                "Input:",  "  Jira Bug & TC file, Template (for Test case output), and root-directory of reports to be updated",
                "Output:", "  Passed reports under directories (named by root-directory-plus-datetime) and TC summary with Linked issues",
            },
            //  "K.Update Report and TC summary (linked issue)",
            new String[] 
            {
                "Update report and TC Linked Issue (linked issue only)", 
                "Input:",  "  Jira Bug & TC file, Template (for Test case output), and root-directory of reports to be updated",
                "Output:", "  Updated reports under directories (named by root-directory-plus-datetime) and TC summary with Linked issues",
            },
            // "L.Create Report, Update Report and TC Summary (A+K)",
            new String[] 
            {
                "Create Report and update them, also update TC Linked Issue ==> Report (A+H)", 
                "Input:",  "  Jira Bug & TC file, Template (for Test case output), and Input Excel File",
                "Output:", "  Updated reports specified within Input Excel File and TC summary with Linked issues",
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
            // "5.Read Report Header to Excel",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Header Template",
            },
            // "6.Write Report Header according to Excel",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "Header Template",
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
            // "B.CST Test Report Creator + update judgement link isssue...etc",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Input Excel File",         // this is file selection
                "TC Template File",
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
            // "J.Update Passed Report and TC summary (linked issue)",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "K.Update All Report and TC summary (linked issue)",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Test Report Path",
                "TC Template File",
            },
            // "L.Create Report, Update Report and TC Summary (A+K)",
            new String[] 
            {
                "Jira Bug File", 
                "Jira TC File",
                "Input Excel File",         // this is file selection
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
            ret_str += "Config file format version: " + ConfigFileFormatVersionString + "\r\n\r\n";
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

        public void LoadConfigAll()
        {
            ConfigFileFormatVersionString = XMLConfig.ReadAppSetting_String("CONFIG_FORMAT_VERSION");
            ConfigFileDefaultValueVersionString = XMLConfig.ReadAppSetting_String("CONFIG_DEFAULT_VALUE_VERSION");

            // config for default filename at MainForm
            this.txtBugFile.Text = XMLConfig.ReadAppSetting_String("workbook_BUG_Jira");
            this.txtTCFile.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Jira");
            this.txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
            this.txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");

            // config for default ExcelAction settings
            ExcelAction.ExcelVisible = XMLConfig.ReadAppSetting_Boolean("Excel_Visible");

            TestCase.LoadFromXML();

            String links = XMLConfig.ReadAppSetting_String("LinkIssueFilterStatusString");
            ReportGenerator.List_of_status_to_filter_for_tc_linked_issue = ReportGenerator.SplitCommaSeparatedStringIntoList(links);

            ReportGenerator.TestReport_Default_Judgement = XMLConfig.ReadAppSetting_String("TestReport_Default_Judgement");
            ReportGenerator.TestReport_Default_Conclusion = XMLConfig.ReadAppSetting_String("TestReport_Default_Conclusion");
            ReportGenerator.TestReport_SaveReportByStatus = XMLConfig.ReadAppSetting_String("TestReport_SaveReportByStatus");
            ReportGenerator.TestReport_ExtraSavePassReport = XMLConfig.ReadAppSetting_Boolean("TestReport_ExtraSavePassReport");
            // End

            // config for default parameters used in Test Plan / Test Report
            TestPlan.NameDefinitionRow_TestPlan = XMLConfig.ReadAppSetting_int("TestPlan_Row_NameDefine");
            TestPlan.DataBeginRow_TestPlan = XMLConfig.ReadAppSetting_int("TestPlan_Row_DataBegin");
            ReportKeyword.col_keyword = XMLConfig.ReadAppSetting_int("KeywordIssue_Column_Keyword_Location");
            ReportKeyword.regexKeywordString = XMLConfig.ReadAppSetting_String("KeywordIssue_Regex_Keyword_Indentifier");
            ReportKeyword.row_offset_result_title = XMLConfig.ReadAppSetting_int("KeywordIssue_Result_Title_Offset_Row");
            ReportKeyword.col_offset_result_title = XMLConfig.ReadAppSetting_int("KeywordIssue_Result_Title_Offset_Col");
            ReportKeyword.row_offset_bugstatus_title = XMLConfig.ReadAppSetting_int("KeywordIssue_BugStatus_Title_Offset_Row");
            ReportKeyword.col_offset_bugstatus_title = XMLConfig.ReadAppSetting_int("KeywordIssue_BugStatus_Title_Offset_Col");
            ReportKeyword.row_offset_buglist_title = XMLConfig.ReadAppSetting_int("KeywordIssue_BugList_Title_Offset_Row");
            ReportKeyword.col_offset_buglist_title = XMLConfig.ReadAppSetting_int("KeywordIssue_BugList_Title_Offset_Col");

            // end of config

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

            ReportGenerator.copy_bug_list = XMLConfig.ReadAppSetting_Boolean("TCSummary_copy_bug_list");
            ReportGenerator.copy_and_extend_bug_list = XMLConfig.ReadAppSetting_Boolean("TCSummary_copy_and_extend_bug_list");
            ReportGenerator.Option_update_status_even_no_report = XMLConfig.ReadAppSetting_Boolean("TCSummary_update_status_even_no_report");

            // Input Excel
            InputExcel.LoadFromXML();

            Issue.LoadFromXML();

            TestReport.LoadFromConfig();

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
        static public String ConfigFileFormatVersionString;
        static public String ConfigFileDefaultValueVersionString;

        private void MainForm_Load(object sender, EventArgs e)
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
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

        /*
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
                // update LUT
                MsgWindow.AppendText("bug_list finished!\n");
                return true;
            }
        }
        */

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
                //return ReadGlobalIssueListTask(filename);
                Boolean bRet = ReportGenerator.ProcessBugListExcel(filename);
                return bRet;
            }
            else
            {
                return true;
            }
        }

        private void ClearIssueList()
        {
            ReportGenerator.ClearGlobalIssueList();
            TestReport.ClearGlobalKeywordList();
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
            TestReport.ClearGlobalKeywordList();
        }

        //private bool Execute_WriteIssueDescriptionToTC_rev2(String tc_file, String template_file, String buglist_file, String judgement_report_dir = "")
        //{
        //    if ((judgement_report_dir != "") && !Storage.DirectoryExists(judgement_report_dir))
        //    {
        //        // if judgement_report_dir has been specified, it must exist
        //        return false;
        //    }

        //    // open bug and process bug
        //    if (ReportGenerator.OpenProcessBugExcel(buglist_file) == false)
        //    {
        //        return false;
        //    }

        //    // open tc and process tc
        //    if (ReportGenerator.OpenProcessTeseCaseExcel(tc_file) == false)
        //    {
        //        return false;
        //    }

        //    // open template and copy bug into it
        //    if (ReportGenerator.OpenTCTemplatePasteBug(template_file) == false)
        //    {
        //        return false;
        //    }

        //    // close bug excel
        //    if (Issue.CloseBugListExcel() == false)
        //    {
        //        return false;
        //    }

        //    // copy tc to template
        //    if (ExcelAction.CopyTestCaseIntoTemplate_v2() == false)
        //    {
        //        MainForm.SystemLogAddLine("Failed @ return of CopyTestCaseIntoTemplate_v2");
        //        return false;
        //    }

        //    if (ReportGenerator.WriteBacktoTCJiraExcelV3_ProcessData(judgement_report_dir: judgement_report_dir) == false)
        //    {
        //        MainForm.SystemLogAddLine("Failed @ return of WriteBacktoTCJiraExcelV3_simpliified_branch_writing_template_by_TC");
        //        return false;
        //    }

        //    // close tc
        //    ExcelAction.CloseTestCaseExcel();

        //    // save tempalte
        //    // 6. Write to another filename with datetime (and close template file)
        //    string dest_filename = Storage.GenerateFilenameWithDateTime(tc_file, FileExt: ".xlsx");
        //    ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename, IsTemplate: true);

        //    return true;
        //}

        /*
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
        */

        private bool Execute_WriteIssueDescriptionToSummary(String template_file)
        {
            if ((ReportGenerator.IsGlobalIssueListEmpty()) || (ReportGenerator.IsGlobalTestcaseListEmpty()) ||
                (!Storage.FileExists(template_file)))
            {
                // protection check
                return false;
            }

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

        /*
        private bool Execute_KeywordIssueGenerationTask(String FileOrDirectoryName, Boolean IsDirectory = false)
        {
            String output_report_path;  // not used for this task
            return Execute_KeywordIssueGenerationTask_returning_report_path(FileOrDirectoryName, IsDirectory, out output_report_path);
        }
        */

        /*
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
                source_dir = FileOrDirectoryName;
            }
            output_report_path = Storage.GenerateDirectoryNameWithDateTime(source_dir);
            TestReport.KeywordIssueGenerationTaskV4(file_list, source_dir, output_report_path);
            return true;
        }
        */

        private Boolean Execute_UpdateReportByLinkedIssue(String FileOrDirectoryName, Boolean IsDirectory, out String output_report_path)
        {
            List<String> file_list = new List<String>();
            String source_dir;
            output_report_path = "";

            if (ReportGenerator.IsGlobalIssueListEmpty())
            {
                // protection check
                return false;
            }

            GetReportFileListForUpdate(FileOrDirectoryName, IsDirectory, out source_dir, out file_list);
            output_report_path = Storage.GenerateDirectoryNameWithDateTime(source_dir);

            TestReport.UpdateReportOnlyByLinkedIssue(file_list, source_dir, output_report_path);
            return true;
        }

        private Boolean Execute_UpdateReportByLinkedIssue_prefiltered_by_tc_status(String FileOrDirectoryName, Boolean IsDirectory, out String output_report_path)
        {
            List<String> file_list = new List<String>();
            String source_dir;
            output_report_path = "";

            if (ReportGenerator.IsGlobalIssueListEmpty())
            {
                // protection check
                return false;
            }

            GetReportFileListForUpdate(FileOrDirectoryName, IsDirectory, out source_dir, out file_list);
            output_report_path = Storage.GenerateDirectoryNameWithDateTime(source_dir);

            List<String> filtered_filelist = ReportGenerator.FilterReportFileListByTCStatus(file_list);

            TestReport.UpdateReportOnlyByLinkedIssue(filtered_filelist, source_dir, output_report_path);
            return true;
        }

        private Boolean GetReportFileListForUpdate(String FileOrDirectory, Boolean IsDirectory, out String RootDir, out List<String> ReportFileList)
        {
            Boolean b_ret = false;

            List<String> AllFileList = new List<String>();
            ReportFileList = AllFileList;
            RootDir = "";

            if (IsDirectory == false)
            {
                // Is a File
                if (Storage.FileExists(FileOrDirectory))
                {
                    String FileName = FileOrDirectory;
                    AllFileList.Add(FileName);
                    RootDir = Storage.GetDirectoryName(FileName);
                    ReportFileList = Storage.FilterFilename(AllFileList);
                    b_ret = true;
                }
            }
            else
            {
                // Is a directory
                if (Storage.DirectoryExists(FileOrDirectory))
                {
                    String DirectoryName = FileOrDirectory;
                    AllFileList = Storage.ListFilesUnderDirectory(DirectoryName);
                    RootDir = DirectoryName;
                    ReportFileList = Storage.FilterFilename(AllFileList);
                    b_ret = true;
                }
            }

            return b_ret;
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

        /*
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
        */

        /*
        private bool Execute_ListAllDetailedTestPlanKeywordTask(String report_root, String output_file = "")
        {
            if (!Storage.DirectoryExists(report_root))
            {
                // protection check
                return false;
            }

            List<TestReportKeyword> keyword_list = TestReport.ListAllDetailedTestPlanKeywordTask(report_root, output_file);

            return true;
        }
        */

        /*
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
*/

        private bool ConfigurableReportUpdate_Task(String excel_input_file)
        {
            if (!Storage.FileExists(excel_input_file))
            {
                // protection check
                return false;
            }

            CopyReport.UpdateTestReportByOptionAndSaveAsAnother(excel_input_file);

            return true;
        }

        private bool Execute_Man_Power_Processing_Task(String excel_input_file)
        {
            if (!Storage.FileExists(excel_input_file))
            {
                // protection check
                return false;
            }

            ManPowerTask.AllHolidayList = ManPowerTask.LoadSiteHolidayList();
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

            TestReport.Update_Group_Summary(report_path);

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
            if (String.IsNullOrWhiteSpace(ret_str) == false)
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
            if (String.IsNullOrWhiteSpace(ret_str) == false)
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
                // In case this field is for selecting file path instaed of directory path
                case ReportType.KeywordIssue_Report_SingleFile:
                case ReportType.FinalCSTReport:                              //Report B   
                case ReportType.Update_Repoart_A_then_Report_K:              //Report L
                    init_dir = Storage.GetFullPath(txtReportFile.Text);
                    sel_file = true;  // Here select file instead of directory
                    break;
                default:
                    init_dir = txtReportFile.Text;
                    break;
            }
            String ret_str = SelectDirectoryOrFile(init_dir, sel_file);
            if (String.IsNullOrWhiteSpace(ret_str) == false)
            {
                txtReportFile.Text = ret_str;
                btnSelectReportFile_Clicked = true;
            }
        }

        private void btnSelectOutputTemplate_Click(object sender, EventArgs e)
        {
            ReportType report_type = Get_comboBoxReportSelect_ReturnReportType();
            bool sel_file = true;
            String init_dir = Storage.GetFullPath(txtOutputTemplate.Text);

            if (report_type == ReportType.Man_Power_Processing)
            {
                String ret_str = Storage.UsesrSelectCSVFilename(init_dir);
                if (String.IsNullOrWhiteSpace(ret_str) == false)
                {
                    txtOutputTemplate.Text = ret_str;
                    XMLConfig.AddUpdateAppSettings("ManPower_last_selected_file", ret_str);
                    //btnSelectOutputTemplate_Clicked = true;
                    btnSelectOutputTemplate_Clicked = false;        // force to always reload after report selection changed.
                }
            }
            else
            {
                String ret_str = SelectDirectoryOrFile(init_dir, sel_file);
                if (String.IsNullOrWhiteSpace(ret_str) == false)
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
        static public List<String> SystemLogPopToListOfString()
        {
            List<String> ret_list = SYSTEM_LOG.ToList();
            SYSTEM_LOG.Clear();
            return ret_list;
        }

        //static private Stack<Boolean> temp_option_stack = new Stack<Boolean>();

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            bool bRet = false;

            ReportType report_type = Get_comboBoxReportSelect_ReturnReportType();
            ClearIssueList();
            ClearTCList();

            UpdateUIDuringExecution(report_type: report_type, executing: true);

            String msg = "Executing: " + GetReportName(report_type) + ".\n";
            MsgWindow.AppendText(msg);
            SystemLogAddLine(msg);

            msg = "Starting time: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ".\n";
            MsgWindow.AppendText(msg);
            SystemLogAddLine(msg);

            Boolean open_excel_ok = false;
            if (report_type != ReportType.Man_Power_Processing)
            {
                open_excel_ok = ExcelAction.OpenExcelApp();
            }
            else
            {
                open_excel_ok = true;
            }
            if (open_excel_ok)
            {
                // Must be updated if new report type added #NewReportType
                switch (report_type)
                {
                    case ReportType.FullIssueDescription_TC:
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        if (ReportGenerator.Process_BugList_TeseCase_TCTemplate(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                        {
                            MainForm.SystemLogAddLine("Failed @ return of Process_BugList_TeseCase_TCTemplate()");
                            bRet = false;
                        }
                        else
                        {
                            bRet = ReportGenerator.Execute_ExtendLinkIssueAndUpdateStatusWithoutReport(tc_file: txtTCFile.Text);
                        }
                        break;
                    case ReportType.FullIssueDescription_Summary: // report 2 not used now
                        /*
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        bRet = Execute_WriteIssueDescriptionToSummary(template_file: txtOutputTemplate.Text);
                        */
                        break;
                    case ReportType.CreateImportToJiraCSV:
                        /*
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        //bRet = Execute_CreateStandardTestReportTask(main_file: txtOutputTemplate.Text);
                        */
                        break;
                    case ReportType.KeywordIssue_Report_SingleFile:                 // Report 4
                        /*
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile) == false) break;  // File path here
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        bRet = Execute_KeywordIssueGenerationTask(FileOrDirectoryName: txtReportFile.Text, IsDirectory: true);
                        */
                        break;
                    case ReportType.KeywordIssue_Report_Directory:                  // Report 7
                        /*
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile) == false) break;  // Directory path here
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        bRet = Execute_KeywordIssueGenerationTask(FileOrDirectoryName: txtReportFile.Text, IsDirectory: true);
                        */
                        break;
                    case ReportType.ReadReportContentVariable:      // Report 5
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        bRet = ReportManagement.ChangeReportHeaderTemplateTask(input_excel_file: txtOutputTemplate.Text);
                        break;
                    case ReportType.WriteReportContentVariable:     // Report 6
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        // to-be-updated
                        //bRet = ConfigurableReportUpdate_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        break;
                    case ReportType.Excel_Sheet_Name_Update_Tool:
                        /*
                        if (UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile) == false) break;  // Directory path here
                        // bRet = Execute_KeywordIssueGenerationTask(txtReportFile.Text, IsDirectory: true);
                        bRet = true;
                        */
                        break;
                    case ReportType.FullIssueDescription_TC_report_judgement:           // Report 9
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        if (ReportGenerator.Process_BugList_TeseCase_TCTemplate(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                        {
                            MainForm.SystemLogAddLine("Failed @ return of OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC()");
                            bRet = false;
                        }
                        else
                        {
                            bRet = ReportGenerator.Execute_UpdateLinkedIssueStatusOnTCTemplate(tc_file: txtTCFile.Text, report_dir: txtReportFile.Text);
                        }
                        break;
                    case ReportType.CreateCSTReport:                                    // Report A
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        TestReport.Report_A_Push_Option();
                        bRet = ConfigurableReportUpdate_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        TestReport.Report_A_Pop_Option();
                        break;
                    case ReportType.FinalCSTReport:                                     // Report B

                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile) == false) break;
                        if (ReportGenerator.Process_BugList_TeseCase_TCTemplate(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                        {
                            MainForm.SystemLogAddLine("Failed @ return of OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC()");
                            bRet = false;
                        }
                        else
                        {
                            String InputExcel = txtReportFile.Text;
                            String ReturnDestinationPaht = "";
                            List<String> DestinationReportList;
                            TestReport.Report_B_Push_Option();
                            bRet = CopyReport.UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(InputExcel, out DestinationReportList, out ReturnDestinationPaht);
                            TestReport.Report_B_Pop_Option();
                            bRet = ReportGenerator.Execute_UpdateLinkedIssueStatusOnTCTemplate(tc_file: txtTCFile.Text, report_list: DestinationReportList);
                        }

                        break;
                    case ReportType.ConfigurableReportUpdate:                          // Report C
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        // to-be-updated
                        bRet = ConfigurableReportUpdate_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        break;
                    case ReportType.CopyReportOnly:                                             // Report D
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        // copy files only
                        TestReport.Report_D_Push_Option();
                        bRet = ConfigurableReportUpdate_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        TestReport.Report_D_Pop_Option();
                        break;
                    case ReportType.RemoveInternalSheet:                                        // Report E
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        // copy files only
                        TestReport.Report_E_Push_Option();
                        bRet = ConfigurableReportUpdate_Task(excel_input_file: Storage.GetFullPath(txtOutputTemplate.Text));
                        TestReport.Report_E_Pop_Option();
                        break;
                    case ReportType.TC_GroupSummaryReport:
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile) == false) break;  // Directory path here
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
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        if (!LoadTCListIfEmpty(txtTCFile.Text)) break;
                        //bRet = Execute_CreateTestReportbyTestCaseTask(report_src_dir: src_dir, output_report_dir: dest_dir);
                        break;
                    case ReportType.Update_Keyword_and_TC_Report: // Report H
                        /*
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        //if (!LoadIssueListIfEmpty(txtBugFile.Text)) break;
                        //if (!LoadTCListIfEmpty(txtTCFile.Text)) break;

                        //bRet = Execute_KeywordIssueGenerationTask_returning_report_path(txtReportFile.Text, true, out report_output_path);
                        //bRet = Execute_WriteIssueDescriptionToTC(tc_file: txtTCFile.Text, judgement_report_dir: report_output_path, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text);
                        if (ReportGenerator.Process_BugList_TeseCase_TCTemplate(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                        {
                            MainForm.SystemLogAddLine("Failed @ return of OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC()");
                            bRet = false;
                        }
                        else
                        {
                            String report_H_output_path;
                            bRet = Execute_KeywordIssueGenerationTask_returning_report_path(txtReportFile.Text, true, out report_H_output_path);
                            bRet = ReportGenerator.Execute_ExtendLinkIssueAndUpdateStatusByReport_v2(tc_file: txtTCFile.Text, report_dir: report_H_output_path);
                        }
                        */
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
                    case ReportType.Update_PassReport_and_TC_Summary:                           //Report J = 1 + filtering out fail report + K (skip TC template update of K)
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        // copied from Report 1
                        if (ReportGenerator.Process_BugList_TeseCase_TCTemplate(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                        {
                            MainForm.SystemLogAddLine("Failed @ return of Process_BugList_TeseCase_TCTemplate()");
                            bRet = false;
                        }
                        else
                        {
                            bRet = ReportGenerator.Execute_ExtendLinkIssueAndUpdateStatusWithoutReport(tc_file: txtTCFile.Text);
                            if (bRet)
                            {
                                String Report_J_return_path;
                                String report_J_input_path = txtReportFile.Text;
                                bRet = Execute_UpdateReportByLinkedIssue_prefiltered_by_tc_status(report_J_input_path, true, out Report_J_return_path);
                            }
                        }
                        break;

                    case ReportType.Update_Report_Linked_Issue_and_TC_Report: // Report K -- report H without keyword function
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxDirToFullAndCheckExist(ref txtReportFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;

                        if (ReportGenerator.Process_BugList_TeseCase_TCTemplate(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                        {
                            MainForm.SystemLogAddLine("Failed @ return of OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC()");
                            bRet = false;
                        }
                        else
                        {
                            String report_output_path_report_k;
                            bRet = Execute_UpdateReportByLinkedIssue(txtReportFile.Text, true, out report_output_path_report_k);
                            if (bRet == true)
                            {
                                bRet = ReportGenerator.Execute_UpdateLinkedIssueStatusOnTCTemplate(tc_file: txtTCFile.Text, report_dir: report_output_path_report_k);
                            }
                        }
                        break;

                    case ReportType.Update_Repoart_A_then_Report_K:                                    //Report L = A + K
                        /*
                                                // copied from report A
                                                // NOTE: Input Excel File is storted in txtReportFile for Report J
                                                if (UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile) == false) break;
                                                String report_l_input_excel_selected = txtReportFile.Text;
                                                String report_l_copied_report_root_path = "";
                                                List<String> report_list_report_l;
                                                Report_A_Push_Option();
                                                bRet = CopyReport.UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(report_l_input_excel_selected, out report_list_report_l, out report_l_copied_report_root_path);
                                                Report_A_Pop_Option();

                                                // copied from report K
                                                if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                                                if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                                                if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;

                                                if (ReportGenerator.OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                                                {
                                                    MainForm.SystemLogAddLine("Failed @ return of OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC()");
                                                    bRet = false;
                                                }
                                                else
                                                {
                                                    String report_L_processed_report_output_path;
                                                    bRet = Execute_KeywordIssueGenerationTask_returning_report_path_simplified(report_l_copied_report_root_path, true, out report_L_processed_report_output_path);
                                                    bRet = ReportGenerator.Execute_ExtendLinkIssueAndUpdateStatusByReport_v2(tc_file: txtTCFile.Text, report_dir: report_L_processed_report_output_path);
                                                }
                        */
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtBugFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtTCFile) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtOutputTemplate) == false) break;
                        if (UpdateTextBoxPathToFullAndCheckExist(ref txtReportFile) == false) break;
                        if (ReportGenerator.Process_BugList_TeseCase_TCTemplate(tc_file: txtTCFile.Text, template_file: txtOutputTemplate.Text, buglist_file: txtBugFile.Text) == false)
                        {
                            MainForm.SystemLogAddLine("Failed @ return of OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC()");
                            bRet = false;
                        }
                        else
                        {
                            String InputExcel = txtReportFile.Text;
                            String ReturnDestinationPaht = "";
                            List<String> DestinationReportList;
                            TestReport.Report_B_Push_Option();
                            TestReport.Option.FunctionC.Copy_Worksheet_AtTheBeginning = false;          // Overridden 
                            TestReport.Option.FunctionC.Copy_Worksheet_AtTheEnd = false;                // Overridden
                            bRet = CopyReport.UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(InputExcel, out DestinationReportList, out ReturnDestinationPaht);
                            TestReport.Report_B_Pop_Option();
                            bRet = ReportGenerator.Execute_UpdateLinkedIssueStatusOnTCTemplate(tc_file: txtTCFile.Text, report_list: DestinationReportList);
                        }

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
            if (report_type != ReportType.Man_Power_Processing)
                ExcelAction.CloseExcelApp();

            MsgWindow.AppendText("Finished.\n");

            msg = "Finished time: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ".\n";
            MsgWindow.AppendText(msg);
            SystemLogAddLine(msg);

            List<String> system_log_strings = SystemLogPopToListOfString();
            foreach (String str in system_log_strings)
            {
                MsgWindow.AppendText(str);
            }

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
                case ReportType.ReadReportContentVariable:          // Report 5
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.WriteReportContentVariable:         // Report 6
                    SetEnable_BugFile(false);
                    SetEnable_TCFile(false);
                    SetEnable_ReportFile(false);
                    SetEnable_OutputTemplate(true);
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
                case ReportType.FinalCSTReport:             // Report B - refer to report J
                    SetEnable_BugFile(true);
                    SetEnable_TCFile(true);
                    SetEnable_ReportFile(true);
                    SetEnable_OutputTemplate(true);
                    break;
                case ReportType.ConfigurableReportUpdate:   // Report C
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
                case ReportType.Update_PassReport_and_TC_Summary:           // Report J 
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
                case ReportType.Update_Repoart_A_then_Report_K:            // Report L = A + K
                    // copied from report J
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
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
                    break;
                case ReportType.ReadReportContentVariable:          // Report 5
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("Report_C_Default_Excel");
                    break;
                case ReportType.WriteReportContentVariable:         // Report 6
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("Report_C_Default_Excel");
                    break;
                case ReportType.Excel_Sheet_Name_Update_Tool:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = @".\SampleData\More chapters_TestCaseID";
                    break;
                case ReportType.FullIssueDescription_TC_report_judgement: // original adopted from "1.Issue Description for TC"
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.CreateCSTReport:                  // report A -- tailored version of report C
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("Report_A_Default_Excel");
                    break;
                case ReportType.FinalCSTReport:                 // Report B - refer to report J
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Report_B_Default_Excel");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.ConfigurableReportUpdate:      // Report C
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
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
                    break;
                case ReportType.Update_Report_Linked_Issue:
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.Update_Keyword_and_TC_Report: // original adopted from report 9     -- report H
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.Man_Power_Processing:
                    //if (!btnSelectOutputTemplate_Clicked)
                    {
                        //String short_userName = Storage.GetWindowsLoginUserName();
                        //txtOutputTemplate.Text = @"C:\Users\" + short_userName + @"\Downloads\Advance Roadmaps.csv";
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("ManPower_last_selected_file");
                        btnSelectOutputTemplate_Clicked = false;
                    }
                    break;
                case ReportType.Update_PassReport_and_TC_Summary:                  // Report J 
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.Update_Report_Linked_Issue_and_TC_Report: // Report K
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("TestReport_default_dir");
                    if (!btnSelectOutputTemplate_Clicked)
                        txtOutputTemplate.Text = XMLConfig.ReadAppSetting_String("workbook_TC_Template");
                    break;
                case ReportType.Update_Repoart_A_then_Report_K:                  // Report L = A + J
                    // copied from report J
                    // NOTE: Input Excel File is storted in txtReportFile for Report L
                    if (!btnSelectReportFile_Clicked)
                        txtReportFile.Text = XMLConfig.ReadAppSetting_String("Report_L_Default_Excel");
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
