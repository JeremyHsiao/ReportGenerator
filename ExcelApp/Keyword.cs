using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Configuration;

namespace ExcelReportApplication
{
    //public enum FindKeywordStatus
    //{
    //    OK = 0,
    //    INIT_STATE,
    //    ERR_OpenDetailExcel_OpenExcelWorkbook,
    //    ERR_OpenDetailExcel_Find_Worksheet,
    //    ERR_CloseDetailExcel_wb_null,
    //    ERR_SaveChangesAndCloseDetailExcel_wb_null,
    //    ERR_NOT_DEFINED,
    //    EX_OpenDetailExcel,
    //    EX_CloseDetailExcel,
    //    EX_SaveChangesAndCloseDetailExcel,
    //    MAX_NO
    //};

    public class TestReportKeyword
    {
        private String keyword;
        private String workbook;
        private String worksheet;
        private int at_row;
        private int at_column;
        private int result_at_row;
        private int result_at_column;
        private int bug_status_at_row;
        private int bug_status_at_column;
        private int bug_list_at_row;
        private int bug_list_at_column;
        private List<Issue> keyword_issues;

        // Generated data where items are according to keyword_list & contents are to-be-defined by requirements.
        private List<String> issue_list;
        private List<String> tc_list;
        private List<StyleString> issue_description_list;
        private List<StyleString> tc_description_list;

        private void TestPlanKeywordInit()
        {
            keyword_issues = new List<Issue>();
            issue_list = new List<String>(); tc_list = new List<String>();
            issue_description_list = new List<StyleString>(); tc_description_list = new List<StyleString>();
        }

        public TestReportKeyword() { TestPlanKeywordInit(); }
        public TestReportKeyword(String Keyword, String Workbook = "", String Worksheet = "", int AtRow = 0, int AtColumn = 0,
                                int ResultListAtRow = 0, int ResultListAtColumn = 0, int BugStatusAtRow = 0, int BugStatusAtColumn = 0,
                                int BugListAtRow = 0, int BugListAtColumn = 0)
        {
            TestPlanKeywordInit();
            keyword = Keyword;
            workbook = Workbook;
            worksheet = Worksheet;
            at_row = AtRow;
            at_column = AtColumn;
            result_at_row = ResultListAtRow;
            result_at_column = ResultListAtColumn;
            bug_status_at_row = BugStatusAtRow;
            bug_status_at_column = BugStatusAtColumn;
            bug_list_at_row = BugListAtRow;
            bug_list_at_column = BugListAtColumn;
        }

        public String Keyword   // property
        {
            get { return keyword; }   // get method
            set { keyword = value; }  // set method
        }

        public String Workbook   // property
        {
            get { return workbook; }   // get method
            set { workbook = value; }  // set method
        }

        public String Worksheet   // property
        {
            get { return worksheet; }   // get method
            set { worksheet = value; }  // set method
        }

        public int AtRow   // property
        {
            get { return at_row; }   // get method
            set { at_row = value; }  // set method
        }
        public int AtColumn   // property
        {
            get { return at_column; }   // get method
            set { at_column = value; }  // set method
        }
        public int ResultAtRow   // propertyd
        {
            get { return result_at_row; }   // get method
            set { result_at_row = value; }  // set method
        }
        public int ResultAtColumn   // property
        {
            get { return result_at_column; }   // get method
            set { result_at_column = value; }  // set method
        }
        public int BugStatusAtRow   // property
        {
            get { return bug_status_at_row; }   // get method
            set { bug_status_at_row = value; }  // set method
        }
        public int BugStatusAtColumn   // property
        {
            get { return bug_status_at_column; }   // get method
            set { bug_status_at_column = value; }  // set method
        }
        public int BugListAtRow   // property
        {
            get { return bug_list_at_row; }   // get method
            set { bug_list_at_row = value; }  // set method
        }
        public int BugListAtColumn   // property
        {
            get { return bug_list_at_column; }   // get method
            set { bug_list_at_column = value; }  // set method
        }

        public List<Issue> KeywordIssues   // property
        {
            get { return keyword_issues; }   // get method
            set { keyword_issues = value; }  // set method
        }

        //public List<String> IssueList   // property
        //{
        //    get { return issue_list; }   // get method
        //    set { issue_list = value; }  // set method
        //}

        public List<String> TestCaseList   // property
        {
            get { return tc_list; }   // get method
            set { tc_list = value; }  // set method
        }

        public List<StyleString> IssueDescriptionList   // property
        {
            get { return issue_description_list; }   // get method
            set { issue_description_list = value; }  // set method
        }

        public List<StyleString> TestCaseDescriptionList   // property
        {
            get { return tc_description_list; }   // get method
            set { tc_description_list = value; }  // set method
        }

        //d
        //public void UpdateIssueDescriptionList(List<StyleString> description)
        //{
        //    List<StyleString> ret_str = new List<StyleString>();

        //    if (IssueList == null)
        //    {
        //        UpdateIssueList();
        //    }

        //    if (IssueList != null)
        //    {

        //    }

        //    IssueDescriptionList = ret_str;
        //}

        //// Greater or Equal to 1.0 ==> not Closed(0) nor Waived(0.1)
        //public IssueCount Calculate_Issue_GE_Stage_1_0()
        //{
        //    IssueCount ret_ic = new IssueCount();
        //    foreach (Issue issue in this.KeywordIssues)
        //    {
        //        if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
        //        {
        //            switch (issue.Severity[0])
        //            {
        //                case 'A':
        //                    ret_ic.Severity_A++;
        //                    break;
        //                case 'B':
        //                    ret_ic.Severity_B++;
        //                    break;
        //                case 'C':
        //                    ret_ic.Severity_C++;
        //                    break;
        //                case 'D':
        //                    ret_ic.Severity_D++;
        //                    break;
        //            }

        //        }
        //    }
        //    return ret_ic;
        //}

        //public IssueCount Calculate_Issue()
        //{
        //    IssueCount ret_ic = new IssueCount();
        //    foreach (Issue issue in this.KeywordIssues)
        //    {
        //        if (issue.Status == Issue.STR_CLOSE)
        //        {
        //            switch (issue.Severity[0])
        //            {
        //                case 'A':
        //                    ret_ic.Closed_A++;
        //                    break;
        //                case 'B':
        //                    ret_ic.Closed_B++;
        //                    break;
        //                case 'C':
        //                    ret_ic.Closed_C++;
        //                    break;
        //                case 'D':
        //                    ret_ic.Closed_D++;
        //                    break;
        //            }
        //        }
        //        else if (issue.Status == Issue.STR_WAIVE)
        //        {
        //            switch (issue.Severity[0])
        //            {
        //                case 'A':
        //                    ret_ic.Waived_A++;
        //                    break;
        //                case 'B':
        //                    ret_ic.Waived_B++;
        //                    break;
        //                case 'C':
        //                    ret_ic.Waived_C++;
        //                    break;
        //                case 'D':
        //                    ret_ic.Waived_D++;
        //                    break;
        //            }
        //        }
        //        else // if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
        //        {
        //            switch (issue.Severity[0])
        //            {
        //                case 'A':
        //                    ret_ic.Severity_A++;
        //                    break;
        //                case 'B':
        //                    ret_ic.Severity_B++;
        //                    break;
        //                case 'C':
        //                    ret_ic.Severity_C++;
        //                    break;
        //                case 'D':
        //                    ret_ic.Severity_D++;
        //                    break;
        //            }

        //        }
        //    }
        //    return ret_ic;
        //}
    }

    public class TestReportBooleanOption
    {
        public Boolean CopyFileOnly = false;
        public Boolean Remove_AUO_Internal = false;
        public Boolean Remove_AUO_Internal_remove_Method = false;
        public Boolean Replace_Conclusion = false;
        public Boolean Update_Report_Sheetname = true;
        public Boolean Clear_Keyword_Result = true;
        public Boolean Hide_Keyword_Result_Bug_Row = false;
        public Boolean Update_Header_by_Template = false;
        public Boolean Update_Conclusion = false;
        public Boolean Update_Judgement = false;
        public Boolean Update_Sample_SN = false;
        // Lagacy options - BEGIN
        public Boolean Report_C_Update_Full_Header = false;
    }

    public class TestReportOption
    {
        public Boolean Report_C_CopyFileOnly = false;
        public Boolean Report_C_Remove_AUO_Internal = false;
        public Boolean Report_C_Remove_AUO_Internal_remove_Method = false; 
        public Boolean Report_C_Replace_Conclusion = false;
        public Boolean Report_C_Update_Report_Sheetname = true;
        public Boolean Report_C_Clear_Keyword_Result = true;
        public Boolean Report_C_Hide_Keyword_Result_Bug_Row = false;
        public Boolean Report_C_Update_Header_by_Template = false;
        public Boolean Report_C_Update_Conclusion = false;
        public Boolean Report_C_Update_Judgement = false;
        public Boolean Report_C_Update_Sample_SN = false;

        public String SampleSN_String = "Refer to DUT_Allocation_Matrix table";
        public String SN_Font = "Gill Sans MT";
        public int SN_FontSize = 12;
        public Color SN_FontColor = Color.Black;
        public FontStyle SN_FontStyle = FontStyle.Regular;
        
        // Lagacy options - BEGIN
        public Boolean Report_C_Update_Full_Header = false;
        public String Report_Title = "Report_Name";
        public String Report_SheetName = "Report_Sheet_Name";
        public String Model_Name = "Model Name";
        public String Part_No = "Part_No";
        public String Panel_Module = "Panel_Module";
        public String TCON_Board = "T-Con_Board";
        public String AD_Board = "AD_Board";
        public String Power_Board = "Power_Board";
        public String Smart_BD_OS_Version = "Smart_BD_OS_Version";
        public String Touch_Sensor = "Touch_Sensor";
        public String Speaker_AQ_Version = "Speaker_AQ_Version";
        public String SW_PQ_Version = "SW_PQ_Version";
        public String Test_Stage = " ";
        public String Test_QTY_SN = " ";
        public String Test_Period_Begin = "2023/07/10";
        public String Test_Period_End = "2023/07/10";
        public String Judgement = " ";
        public String Tested_by = " ";
        public String Approved_by = "Jeremy Hsiao";
        public Boolean Update_Report_Title_by_Sheetname = true;
        public Boolean Update_Model_Name = true;
        public Boolean Update_Part_No = true;
        public Boolean Update_Panel_Module = true;
        public Boolean Update_TCON_Board = true;
        public Boolean Update_AD_Board = true;
        public Boolean Update_Power_Board = true;
        public Boolean Update_Smart_BD_OS_Version = true;
        public Boolean Update_Touch_Sensor = true;
        public Boolean Update_Speaker_AQ_Version = true;
        public Boolean Update_SW_PQ_Version = true;
        public Boolean Update_Test_Stage = true;
        public Boolean Update_Test_QTY_SN = true;
        public Boolean Update_Test_Period_Begin = true;
        public Boolean Update_Test_Period_End = true;
        public Boolean Update_Judgement = true;
        public Boolean Update_Tested_by = true;
        public Boolean Update_Approved_by = true;
        // Lagacy options - END
    }

    public static class TestReport
    {
        static public String TestReport_Default_Output_Dir = "";

        // From TestReportOption - BEGIN
        public static int Title_at_row = 1, Title_at_col = ExcelAction.ColumnNameToNumber('A');

        public static int Period_Start_at_row = 8, Period_Start_at_col = ExcelAction.ColumnNameToNumber('L');
        public static int Period_End_at_row = 8, Period_End_at_col = ExcelAction.ColumnNameToNumber('M');
        //        public static int Judgement_at_row = 9, Judgement_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int Judgement_string_at_row = 9, Judgement_string_at_col = ExcelAction.ColumnNameToNumber('B');

        public static int Model_Name_at_row = 3, Model_Name_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int Part_No_at_row = 3, Part_No_at_col = ExcelAction.ColumnNameToNumber('J');

        public static int Panel_Module_at_row = 4, Panel_Module_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int TCON_Board_at_row = 4, TCON_Board_at_col = ExcelAction.ColumnNameToNumber('J');

        public static int AD_Board_at_row = 5, AD_Board_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int Power_Board_at_row = 5, Power_Board_at_col = ExcelAction.ColumnNameToNumber('J');

        public static int Smart_BD_OS_Version_at_row = 6, Smart_BD_OS_Version_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int Touch_Sensor_at_row = 6, Touch_Sensor_at_col = ExcelAction.ColumnNameToNumber('J');

        public static int Speaker_AQ_Version_at_row = 7, Speaker_AQ_Version_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int SW_PQ_Version_at_row = 7, SW_PQ_Version_at_col = ExcelAction.ColumnNameToNumber('J');

        public static int Test_Stage_at_row = 8, Test_Stage_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int Test_QTY_SN_at_row = 8, Test_QTY_SN_at_col = ExcelAction.ColumnNameToNumber('H');
        public static int Test_Period_Begin_at_row = 8, Test_Period_Begin_at_col = ExcelAction.ColumnNameToNumber('L');
        public static int Test_Period_End_at_row = 8, Test_Period_End_at_col = ExcelAction.ColumnNameToNumber('M');

        public static int Judgement_at_row = 9, Judgement_at_col = ExcelAction.ColumnNameToNumber('D');
        public static int Tested_by_at_row = 9, Tested_by_at_col = ExcelAction.ColumnNameToNumber('H');
        public static int Approved_by_at_row = 9, Approved_by_at_col = ExcelAction.ColumnNameToNumber('L');

        //public static int Part_No_at_row = 3, Part_No_at_col = ExcelAction.ColumnNameToNumber('J');
        //public static int SW_Version_at_row = 7, SW_Version_at_col = ExcelAction.ColumnNameToNumber('J');
        //public static int Period_Start_at_row = 8, Period_Start_at_col = ExcelAction.ColumnNameToNumber('L');
        //public static int Period_End_at_row = 8, Period_End_at_col = ExcelAction.ColumnNameToNumber('M');
        //public static int Judgement_at_row = 9, Judgement_at_col = ExcelAction.ColumnNameToNumber('D');
        //public static int Judgement_string_at_row = 9, Judgement_string_at_col = 2;
        private static StyleString blank_space = new StyleString(" ", ReportGenerator.LinkIssue_report_FontColor,
                            ReportGenerator.LinkIssue_report_Font, ReportGenerator.LinkIssue_report_FontSize);
        public static List<StyleString> blank_space_list = blank_space.ConvertToList();
        
        // From TestReportOption - END

        public static int col_indentifier = ExcelAction.ColumnNameToNumber('B');
        public static int col_keyword = col_indentifier + 1;
        public static int row_test_brief_start = 10;
        public static int row_test_brief_end = 22;
        public static int row_default_conclusion_title = 21;
        public static int row_test_detail_start = 27;
        public static int col_default_report_right_border = ExcelAction.ColumnNameToNumber('N');
        public static String regexKeywordString = @"(?i)Item";
        public static int row_offset_result_title = 1;                                // offset from the row of identifier regex "Item"
        public static int col_offset_result_title = 1;                                // offset from the column of identifier regex "Item"
        public static int row_offset_bugstatus_title = row_offset_result_title;       // offset from the row of identifier regex "Item"
        public static int col_offset_bugstatus_title = col_offset_result_title + 2;   // offset from the column of identifier regex "Item"
        public static int row_offset_buglist_title = row_offset_result_title + 1;     // offset from the row of identifier regex "Item"
        public static int col_offset_buglist_title = col_offset_result_title;         // offset from the column of identifier regex "Item"
        public static String regexResultString = @"^(?i)\s*Result\s*$";
        public static String regexBugStatusString = @"^(?i)\s*Bug Status\s*$";
        public static String regexBugListString = @"^(?i)\s*Bug List\s*$";

        static public List<String> KeywordIssue_filter_status_list = new List<String>();
        public static String TestReport_Default_Judgement = "N/A";
        public static String TestReport_Default_Conclusion = " ";

        public static String PASS_str = "Pass";
        public static String CONDITIONAL_PASS_str = "Conditional Pass";
        public static String FAIL_str = "Fail";
        public static String WAIVED_str = "Waived";

        public static Boolean KeywordIssue_Replace_Conclusion = false;
        public static Boolean KeywordIssue_Hide_Result_Bug = false;

        public static Boolean KeywordIssue_Auto_Correct_Sheetname = false;

        public static TestReportOption Option = new TestReportOption();

        public static int PassCnt_at_row = 21, PassCnt_at_col = ExcelAction.ColumnNameToNumber('E');
        public static int FailCnt_at_row = 21, FailCnt_at_col = ExcelAction.ColumnNameToNumber('G');
        //public static int TotalCnt_at_row = 21, TotalCnt_at_col = 9;
        public static int ConditionalPass_string_at_row = 21, ConditionalPass_string_at_col = ExcelAction.ColumnNameToNumber('H');
        public static int ConditionalPassCnt_at_row = 21, ConditionalPassCnt_at_col = ExcelAction.ColumnNameToNumber('I');


        private static List<TestReportKeyword> global_keyword_list = new List<TestReportKeyword>();
        private static Boolean global_keyword_available;
        public static Boolean CheckGlobalKeywordListExist()
        {
            return global_keyword_available;
        }
        public static void SetGlobalKeywordList(List<TestReportKeyword> keyword_list)
        {
            global_keyword_list = keyword_list;
            global_keyword_available = true;
        }
        public static List<TestReportKeyword> GetGlobalKeywordList()
        {
            if (global_keyword_available)
            {
                return global_keyword_list;
            }
            else
            {
                return new List<TestReportKeyword>();
            }
        }
        public static void ClearGlobalKeywordList()
        {
            global_keyword_available = false;
            global_keyword_list.Clear();
        }

        //private static Dictionary<String, String> global_report_judgement_result = new Dictionary<String, String>();
        //public static Boolean CheckLookupReportJudgementResultExist()
        //{
        //    return (global_report_judgement_result.Count > 0);
        //}
        //public static String LookupReportJudgementResult(String full_report_path)
        //{
        //    String ret_str = "";
        //    if (global_report_judgement_result.ContainsKey(full_report_path))
        //    {
        //        ret_str = global_report_judgement_result[full_report_path];
        //    }
        //    return ret_str;
        //}
        //public static void ClearReportJudgementResult()
        //{
        //    global_report_judgement_result.Clear();
        //}
        //public static void AppendReportJudgementResult(String full_report_path, String judgement)
        //{
        //    global_report_judgement_result.Add(full_report_path, judgement);
        //}

        private enum REPORT_INFO
        {
            JUDGEMENT = 0,
            PURPOSE,
            CRITERIA,
        }
        private static Dictionary<String, List<String>> global_report_information = new Dictionary<String, List<String>>();
        public static Boolean CheckLookupReportInformationExist()
        {
            return (global_report_information.Count > 0);
        }
        public static void ClearReportInformation()
        {
            global_report_information.Clear();
        }
        public static void AppendReportInformation(String full_report_path, List<String> info_to_append)
        {
            global_report_information.Add(full_report_path, info_to_append);
        }
        public static List<String> LookupReportInformation(String full_report_path)
        {
            List<String> ret_list_str = new List<String>();
            if (global_report_information.ContainsKey(full_report_path))
            {
                ret_list_str = global_report_information[full_report_path];
            }
            return ret_list_str;
        }
        public static List<String> CombineReportInfo(String judgement = "", String purpose = "", String criteria = "")
        {
            List<String> ret_list_str = new List<String>();
            ret_list_str.Add(judgement);
            ret_list_str.Add(purpose);
            ret_list_str.Add(criteria);
            return ret_list_str;
        }
        public static String GetJudgement(List<String> info)
        {
            return info.ElementAt((int)REPORT_INFO.JUDGEMENT);
        }
        public static String GetPurpose(List<String> info)
        {
            return info.ElementAt((int)REPORT_INFO.PURPOSE);
        }
        public static String GetCriteria(List<String> info)
        {
            return info.ElementAt((int)REPORT_INFO.CRITERIA);
        }

        // Visit report content and find out all keywords
        static public List<TestReportKeyword> ListKeyword_SingleReport(TestPlan plan)
        {
            //
            // 2. Find out Printable Area
            //
            // Assummed that Printable area always starting at $A$1 (also data processing area)
            // So excel data processing area ends at Printable area (row_count,col_count)
            Worksheet ws_testplan = plan.TestPlanWorksheet;
            //Range rngProcessedRange = ExcelAction.GetWorksheetPrintableRange(ws_testplan);
            Range rngProcessedRange = ExcelAction.GetWorksheetAllRange(ws_testplan);
            int row_end = ExcelAction.Get_Range_RowNumber(rngProcessedRange);
            int col_end = ExcelAction.Get_Range_ColumnNumber(rngProcessedRange);

            //
            // 3. Find out all keywords and create LUT (keyword,row_index)
            //    output:  LUT (keyword,row_index)
            //
            // Read report file for keyword & its row and store into keyword/row dictionary
            // Search keyword within printable area
            Dictionary<String, int> KeywordAtRow = new Dictionary<String, int>();
            RegexStringValidator identifier_keyword_Regex = new RegexStringValidator(regexKeywordString);
            //RegexStringValidator result_keyword_Regex = new RegexStringValidator(regexResultString);
            //RegexStringValidator bug_status_keyword_Regex = new RegexStringValidator(regexBugStatusString);
            //RegexStringValidator bug_list_keyword_Regex = new RegexStringValidator(regexBugListString);
            for (int row_index = row_test_detail_start; row_index <= row_end; row_index++)
            {
                String identifier_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, col_indentifier),
                    //result_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index + row_offset_result_title,
                    //                                                    col_indentifier + col_offset_result_title),
                    //bug_status_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index + row_offset_bugstatus_title,
                    //                                                    col_indentifier + col_offset_bugstatus_title),
                    //bug_list_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index + row_offset_buglist_title,
                    //                                                    col_indentifier + col_offset_buglist_title),
                        keyword_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, col_keyword);
                int regex_step = 0;
                try
                {
                    // Attempt validation.
                    // regex false (not a keyword row) then jumping to catch(); 
                    // if special string found at (row_index, col_indentifier)
                    identifier_keyword_Regex.Validate(identifier_text); regex_step++;
                    // regex true, next step is to check the rest of field to validate
                    // 1. Check "Result" title
                    //result_keyword_Regex.Validate(result_text); regex_step++;
                    //bug_status_keyword_Regex.Validate(bug_status_text); regex_step++;
                    //bug_list_keyword_Regex.Validate(bug_list_text); regex_step++;

                    //if (keyword_text == "") { ConsoleWarning("Empty Keyword", row_index); continue; }
                    if (String.IsNullOrWhiteSpace(keyword_text)) { LogMessage.WriteLine("Empty Keyword at row: " + row_index.ToString()); continue; }
                    if (KeywordAtRow.ContainsKey(keyword_text)) { LogMessage.WriteLine("Duplicated Keyword:" + keyword_text + " at " + row_index.ToString()); continue; }
                    KeywordAtRow.Add(keyword_text, row_index);
                }
                catch (ArgumentException ex)
                {
                    // Validation failed.
                    // Not a key row
                    switch (regex_step)
                    {
                        case 0:
                            // Not a keyword identifier (string beginning with "item")
                            break;
                        case 1:
                            // Not a "Result" 
                            LogMessage.CheckFunctionAtRowColumn(regexBugStatusString, row_index + row_offset_result_title,
                                                                 col_indentifier + col_offset_result_title);
                            break;
                        case 2:
                            // Not a "Bug Status" 
                            LogMessage.CheckFunctionAtRowColumn(regexBugStatusString, row_index + row_index + row_offset_bugstatus_title,
                                                                 col_indentifier + col_offset_bugstatus_title);
                            break;
                        case 3:
                            // Not a "Bug List" 
                            LogMessage.CheckFunctionAtRowColumn(regexBugListString, row_index + row_offset_buglist_title,
                                                                 col_indentifier + col_offset_buglist_title);
                            break;
                        default:
                            break;
                    }
                    continue;
                }
            }

            List<TestReportKeyword> ret = new List<TestReportKeyword>();
            foreach (String key in KeywordAtRow.Keys)
            {
                int row_keyword = KeywordAtRow[key];
                // col_keyword is currently fixed value
                int row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list;

                row_result = row_bug_status = row_keyword + 1;
                row_bug_list = row_keyword + 2;
                col_result = col_bug_list = col_keyword + 1;
                col_bug_status = col_keyword + 3;
                ret.Add(new TestReportKeyword(key, plan.ExcelFile, plan.ExcelSheet, KeywordAtRow[key], col_keyword,
                    row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list));
            }
            return ret;
        }

        // Visit all reports and find out all keywords -- make use of ListKeyword_SingleReport()
        static public List<TestReportKeyword> ListAllKeyword(List<TestPlan> DoPlan)
        {
            List<TestReportKeyword> ret = new List<TestReportKeyword>();
            List<ReportFileRecord> ret_not_report_log = new List<ReportFileRecord>();

            foreach (TestPlan plan in DoPlan)
            {
                String path = Storage.GetDirectoryName(plan.ExcelFile);
                String filename = Storage.GetFileName(plan.ExcelFile);
                String sheet_name = plan.ExcelSheet;
                ReportFileRecord fail_log = new ReportFileRecord(path, filename, sheet_name);

                TestPlan.ExcelStatus test_plan_status;
                test_plan_status = plan.OpenDetailExcel();
                if (test_plan_status == TestPlan.ExcelStatus.OK)
                {
                    List<TestReportKeyword> plan_keyword = ListKeyword_SingleReport(plan);
                    plan.CloseDetailExcel();
                    if (plan_keyword != null)
                    {
                        if (plan_keyword.Count() > 0)
                        {
                            ret.AddRange(plan_keyword);
                            fail_log.SetFlagOK(excelfilenameOK: true, openfileOK: true, findWorksheetOK: true, findAnyKeyword: true);
                            // not adding ok report log at the moment
                            //ret_not_report_log.Add(fail_log);
                        }
                        else
                        {
                            fail_log.SetFlagOK(excelfilenameOK: true, openfileOK: true, findWorksheetOK: true);
                            fail_log.SetFlagFail(findNoKeyword: true);
                            ret_not_report_log.Add(fail_log);
                        }
                    }
                    else // (null) 
                    {
                        fail_log.SetFlagOK(excelfilenameOK: true, openfileOK: true, findWorksheetOK: true);
                        fail_log.SetFlagFail(findNoKeyword: true, otherFailure: true);
                        ret_not_report_log.Add(fail_log);
                        LogMessage.WriteLine("Test Plan null keyword list Error occurred:" + plan.ExcelSheet + "@" + plan.ExcelFile);
                    }
                }
                else
                {
                    if (test_plan_status == TestPlan.ExcelStatus.ERR_OpenDetailExcel_OpenExcelWorkbook)
                    {
                        fail_log.SetFlagFail(openfileFail: true);
                    }
                    else if (test_plan_status == TestPlan.ExcelStatus.ERR_OpenDetailExcel_Find_Worksheet)
                    {
                        fail_log.SetFlagOK(excelfilenameOK: true, openfileOK: true);
                        fail_log.SetFlagFail(findWorksheetFail: true);
                    }
                    else
                    {
                        fail_log.SetFlagFail(openfileFail: true, otherFailure: true);
                        LogMessage.WriteLine("Test Plan Unknown Error occurred:" + plan.ExcelSheet + "@" + plan.ExcelFile);
                    }
                    ret_not_report_log.Add(fail_log);
                    plan.CloseDetailExcel();
                }
            }
            ReportGenerator.excel_not_report_log.AddRange(ret_not_report_log);
            return ret;
        }

        //static public List<TestPlanKeyword> FilterSingleReportKeyword(List<TestPlanKeyword> keyword_list, String workbook, String worksheet)
        //{
        //    List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
        //    foreach (TestPlanKeyword kw in keyword_list)
        //    {
        //        if (((kw.Workbook == workbook) || (workbook == "")) && (kw.Worksheet == worksheet))
        //        {
        //            ret.Add(kw);
        //        }
        //    }
        //    return ret;
        //}

        //static public List<TestPlanKeyword> FilterSingleReportKeyword_check_only_worksheet(List<TestPlanKeyword> keyword_list, String worksheet)
        //{
        //    List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
        //    foreach (TestPlanKeyword kw in keyword_list)
        //    {
        //        if ((kw.Worksheet == worksheet))
        //        {
        //            ret.Add(kw);
        //        }
        //    }
        //    return ret;
        //}

        //static public Boolean IsAnyKeywordInReport(List<TestPlanKeyword> keyword_list, String workbook, String worksheet)
        //{
        //    Boolean b_ret = false;
        //    foreach (TestPlanKeyword kw in keyword_list)
        //    {
        //        if ((kw.Workbook == workbook) && (kw.Worksheet == worksheet))
        //        {
        //            b_ret = true;
        //            break;
        //        }
        //    }
        //    return b_ret;
        //}

        // List all duplicated keywords based on complete keyword-list
        static public List<TestReportKeyword> ListAllDuplicatedKeyword(List<TestReportKeyword> keyword_list)
        {
            Dictionary<String, TestReportKeyword> for_checking_duplicated = new Dictionary<String, TestReportKeyword>();
            Dictionary<String, List<TestReportKeyword>> dic_ret_kw_list = new Dictionary<String, List<TestReportKeyword>>();

            foreach (TestReportKeyword keyword in keyword_list)
            {
                String kw = keyword.Keyword;
                if (for_checking_duplicated.ContainsKey(kw))
                {
                    // found duplicated item ==> kw exists in for_checking_duplicated
                    // first to check if already duplicated before
                    if (dic_ret_kw_list.ContainsKey(kw) == false)
                    {
                        // 1st time duplicated so that not available in dic_ret_kw_list
                        // then it is necessary to create a new item in dic_ret_kw_list
                        List<TestReportKeyword> new_duplicated_list = new List<TestReportKeyword>();
                        new_duplicated_list.Add(for_checking_duplicated[kw]);
                        dic_ret_kw_list.Add(kw, new_duplicated_list);
                    }
                    // kw exists in dic_ret_kw_list
                    //add duplicated item into dic_ret_kw_list[kw]
                    dic_ret_kw_list[kw].Add(keyword);
                }
                else
                {
                    // not duplicated item
                    // add this item into for_checking_duplicated[kw]
                    for_checking_duplicated.Add(kw, keyword);
                }
            }

            List<TestReportKeyword> ret_dup_kw_list = new List<TestReportKeyword>();
            foreach (String kw in dic_ret_kw_list.Keys)
            {
                ret_dup_kw_list.AddRange(dic_ret_kw_list[kw]);
            }
            return ret_dup_kw_list;
        }

        // List all duplicated keywords based on complete keyword-list
        static public List<String> ListDuplicatedKeywordString(List<TestReportKeyword> keyword_list)
        {
            SortedSet<String> check_duplicated_keyword = new SortedSet<String>();
            List<TestReportKeyword> duplicate_keyword_list = ListAllDuplicatedKeyword(keyword_list);
            foreach (TestReportKeyword keyword in duplicate_keyword_list)
            {
                check_duplicated_keyword.Add(keyword.Keyword);
            }
            List<String> ret_str_list = new List<String>();
            ret_str_list.AddRange(check_duplicated_keyword);
            return ret_str_list;
        }

        //
        // This Demo is to identify Keyword on the excel and insert a column to list all issues containing that keyword
        //
        //static int col_indentifier = 2;
        //static int col_keyword = 3;
        //        static public bool KeywordIssueGenerationTask(string report_filename)
        /*
                {
                    //
                    // 1. Open Excel and find the sheet
                    //

                    String full_filename = Storage.GetFullPath(report_filename);
                    String short_filename = Storage.GetFileName(full_filename);
                    String sheet_name = short_filename.Substring(0, short_filename.IndexOf("_"));

                    // File exist check is done outside

                    Workbook wb_keyword_issue = ExcelAction.OpenExcelWorkbook(full_filename, ReadOnly: false);
                    if (wb_keyword_issue == null)
                    {
                        ConsoleWarning("OpenExcelWorkbook in KeywordIssueGenerationTask");
                        return false;
                    }

                    Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_keyword_issue, sheet_name);
                    if (result_worksheet == null)
                    {
                        ConsoleWarning("Find_Worksheet in KeywordIssueGenerationTask");
                        return false;
                    }

                    //
                    // 2. Find out Printable Area
                    //
                    // Assummed that Printable area always starting at $A$1 (also data processing area)
                    // So excel data processing area ends at Printable area (row_count,col_count)
                    Range rngPrintable = ExcelAction.GetWorksheetPrintableRange(result_worksheet);
                    int row_print_area = rngPrintable.Rows.Count;
                    int column_print_area = rngPrintable.Columns.Count;

                    //
                    // 3. Find out all keywords and create LUT (keyword,row_index)
                    //    output:  LUT (keyword,row_index)
                    //
                    const int row_test_detail_start = 27;
                    const String identifier_str = "Item";
                    // Read report file for keyword & its row and store into keyword/row dictionary
                    // Search keyword within printable area
                    Dictionary<String, int> KeywordAtRow = new Dictionary<String, int>();
                    for (int row_index = row_test_detail_start; row_index <= row_print_area; row_index++)
                    {
                        String cell_text = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, col_indentifier);
                        if (cell_text == "") continue;
                        if ((cell_text.Length > identifier_str.Length) &&
                            (cell_text.ToLowerInvariant().Contains(identifier_str.ToLowerInvariant())))
                        {
                            cell_text = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, col_keyword);
                            if (cell_text == "") { ConsoleWarning("Empty Keyword", row_index); continue; }
                            if (KeywordAtRow.ContainsKey(cell_text)) { ConsoleWarning("Duplicated Keyword", row_index); continue; }
                            KeywordAtRow.Add(cell_text, row_index);
                        }
                    }

                    //
                    // 4. Use keyword to find out all issues that contains keyword. 
                    //    put issue_id into a string contains many id separated by a comma ','
                    //    then store this issue_id into LUT (keyword,ids)
                    //    output: LUT (keyword,id_list)
                    //
                    Dictionary<String, List<String>> KeywordIssueIDList = new Dictionary<String, List<String>>();
                    foreach (String keyword in KeywordAtRow.Keys)
                    {
                        List<String> id_list = new List<String>();
                        foreach (Issue issue in ReportGenerator.global_issue_list)
                        {
                            if (issue.ContainKeyword(keyword))
                            {
                                id_list.Add(issue.Key);
                            }
                        }
                        KeywordIssueIDList.Add(keyword, id_list);
                    }

                    //
                    // 5. input:  LUT (keyword,id_list) + LUT (id,color_desription) (from GenerateIssueDescription())
                    //    output: LUT (keyword,color_desription_list)
                    //         
                    //    using: id_list -> ExtendIssueDescription() -> color_description_list
                    // This issue description list is needed for keyword issue list
                    ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);

                    // Go throught each keyword and turn id_list into color_description
                    Dictionary<String, List<StyleString>> KeyWordIssueDescription = new Dictionary<String, List<StyleString>>();
                    foreach (String keyword in KeywordAtRow.Keys)
                    {
                        List<String> id_list = KeywordIssueIDList[keyword];
                        List<StyleString> issue_description = StyleString.ExtendIssueDescription(id_list, ReportGenerator.global_issue_description_list);
                        KeyWordIssueDescription.Add(keyword, issue_description);
                    }

                    //
                    // 6. input:  LUT (keyword,color_description_list) + LUT (id,row_index)
                    //    output: write color_description_list at Excel(row_index,new_inserted_col outside printable area
                    //         
                    // Insert extra column just outside printable area.
                    int insert_col = column_print_area + 1;
                    ExcelAction.Insert_Column(result_worksheet, insert_col);

                    foreach (String keyword in KeywordAtRow.Keys)
                    {
                        List<StyleString> issue_description = KeyWordIssueDescription[keyword];
                        StyleString.WriteStyleString(result_worksheet, KeywordAtRow[keyword], insert_col, issue_description);
                    }

                    // Save as another file with yyyyMMddHHmmss
                    string dest_filename = Storage.GenerateFilenameWithDateTime(full_filename);
                    ExcelAction.CloseExcelWorkbook(wb_keyword_issue, SaveChanges: true, AsFilename: dest_filename);
                    return true;
                }
        */
        /*
                static public bool KeywordIssueGenerationTaskV2(string report_filename)
                {
                    //
                    // 1. Find keyword for user selected file
                    //
                    String full_filename = Storage.GetFullPath(report_filename);
                    String short_filename = Storage.GetFileName(full_filename);
                    String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                    String sheet_name = sp_str[0];
                    String subpart = sp_str[1];

                    // Create a temporary test plan -- DoOrNot must be "V" & ExcelFile/ExcelSheet must be correct
                    List<String> tp_str = new List<String>();
                    tp_str.AddRange(new String[] { "N/A", short_filename, "N/A", "V", "N/A", subpart });
                    TestPlan tp = new TestPlan(tp_str);
                    tp.ExcelFile = full_filename;
                    tp.ExcelSheet = sheet_name;
                    List<TestPlan> do_plan = new List<TestPlan>();
                    do_plan.Add(tp);

                    // List all keyword within this temprary test plan
                    List<TestPlanKeyword> keyword_list = KeywordReport.ListAllKeyword(do_plan);

                    // 2. Open Excel and find the sheet
                    // File exist check is done outside
                    Workbook wb_keyword_issue = ExcelAction.OpenExcelWorkbook(full_filename);
                    if (wb_keyword_issue == null)
                    {
                        ConsoleWarning("OpenExcelWorkbook in KeywordIssueGenerationTaskV2");
                        return false;
                    }

                    Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_keyword_issue, sheet_name);
                    if (result_worksheet == null)
                    {
                        ConsoleWarning("Find_Worksheet in KeywordIssueGenerationTaskV2");
                        return false;
                    }

                    //
                    // 3. Use keyword to find out all issues (ID) that contains keyword on id_list. 
                    //    Extend list of issue ID to list of issue description (with font style settings)
                    //
                    ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);
                    Dictionary<String, List<String>> KeywordIssueIDList = new Dictionary<String, List<String>>();
                    foreach (Issue issue in ReportGenerator.global_issue_list)
                    {
                        issue.KeywordList.Clear();
                    }
                    foreach (TestPlanKeyword keyword in keyword_list)
                    {
                        List<StyleString> description_list;
                        List<String> id_list = new List<String>();
                        String keyword_str = keyword.Keyword;
                        foreach (Issue issue in ReportGenerator.global_issue_list)
                        {
                            if (issue.ContainKeyword(keyword_str))
                            {
                                id_list.Add(issue.Key);
                                issue.KeywordList.Add(keyword_str);
                            }
                        }
                        keyword.IssueList = id_list;
                        description_list = StyleString.ExtendIssueDescription(id_list, ReportGenerator.global_issue_description_list);
                        keyword.IssueDescriptionList = description_list;
                    }

                    //
                    // 4. input:  IssueDescriptionList of Keyword
                    //    output: write color_description_list at Excel(row_index,new_inserted_col outside printable area
                    //         
                    // Insert extra column just outside printable area.
                    // Assummed that Printable area always starting at $A$1 (also data processing area)
                    // So excel data processing area ends at Printable area (row_count,col_count)
                    int column_print_area = ExcelAction.GetWorksheetPrintableRange(result_worksheet).Columns.Count;
                    int insert_col = column_print_area + 1;
                    ExcelAction.Insert_Column(result_worksheet, insert_col);

                    foreach (TestPlanKeyword keyword in keyword_list)
                    {
                        int at_row = keyword.AtRow;
                        StyleString.WriteStyleString(result_worksheet, at_row, insert_col, keyword.IssueDescriptionList);
                    }

                    // Save as another file with yyyyMMddHHmmss
                    string dest_filename = Storage.GenerateFilenameWithDateTime(full_filename);
                    ExcelAction.CloseExcelWorkbook(wb_keyword_issue, SaveChanges: true, AsFilename: dest_filename);
                    return true;
                }
        */
        /*
                static public bool KeywordIssueGenerationTaskV3(List<String> report_filename)
                {
                    //
                    // 1. Create a temporary test plan (do_plan) to include all report files 
                    //
                    // 1.1 Init an empty plan
                    List<TestPlan> do_plan = new List<TestPlan>();

                    // 1.2 This temporary test plan starts to includes all files listed in List<String> report_filename
                    foreach(String name in report_filename)
                    {
                        // File existing check protection (it is better also checked and giving warning before entering this function)
                         if (Storage.FileExists(name)==false)
                            continue; // no warning here, simply skip this file.
            
                        // DoOrNot must be "V" & ExcelFile/ExcelSheet must be correct
                        String full_filename = Storage.GetFullPath(name);
                        String short_filename = Storage.GetFileName(full_filename);
                        String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                        String sheet_name = sp_str[0];
                        String subpart = sp_str[1];
                        List<String> tp_str = new List<String>();
                        tp_str.AddRange(new String[] { "N/A", short_filename, "N/A", "V", "N/A", subpart });
                        TestPlan tp = new TestPlan(tp_str);
                        tp.ExcelFile = full_filename;
                        tp.ExcelSheet = sheet_name;
                        do_plan.Add(tp);
                    }

                    //
                    // 2. Search keywords within all selected file (2.1) and use those keywords to find out issues containing keywords.
                   //
                    // 2.1. Find keyword for all selected file (as listed in temprary test plan)
                    //
                    List<TestPlanKeyword> keyword_list = KeywordReport.ListAllKeyword(do_plan);

                    //
                    // 2.2. Use keyword to find out all issues (ID) that contains keyword on id_list. 
                    //    Extend list of issue ID to list of issue description (with font style settings) -- by Issue.GenerateIssueDescription
                    //
                    //ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);//done outside in advance
                    Dictionary<String, List<String>> KeywordIssueIDList = new Dictionary<String, List<String>>();
                    foreach (Issue issue in ReportGenerator.global_issue_list)
                    {
                        issue.KeywordList.Clear();
                    }
                    // Go throught each keyword, search all issues containing this keyword and add issue-id so that it can be extened into description list.
                    foreach (TestPlanKeyword keyword in keyword_list)
                    {
                        List<StyleString> description_list;
                        List<String> id_list = new List<String>();
                        String keyword_str = keyword.Keyword;
                        foreach (Issue issue in ReportGenerator.global_issue_list)
                        {
                            if (issue.ContainKeyword(keyword_str))
                            {
                                id_list.Add(issue.Key);
                                issue.KeywordList.Add(keyword_str);
                            }
                        }
                        keyword.IssueList = id_list;
                        description_list = StyleString.ExtendIssueDescription(id_list, ReportGenerator.global_issue_description_list);
                        keyword.IssueDescriptionList = description_list;
                    }

                    //
                    // 3. Go throught each report excel and generate keyword report for each one.
                    //
                    foreach (TestPlan plan in do_plan)
                    {
                        String full_filename = plan.ExcelFile;
                        String sheet_name = plan.ExcelSheet;

                        // 3.1. Open Excel and find the sheet
                        // File exist check is done outside
                        Workbook wb_keyword_issue = ExcelAction.OpenExcelWorkbook(full_filename);
                        if (wb_keyword_issue == null)
                        {
                            ConsoleWarning("ERR: Open workbook in V3: " + full_filename);
                            return false;
                        }

                        // 3.2 Open worksheet
                        Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_keyword_issue, sheet_name);
                        if (result_worksheet == null)
                        {
                            ConsoleWarning("ERR: Open worksheet in V3: " + full_filename + " sheet: " + sheet_name);
                            return false;
                        }

                        //
                        // 3.3. input:  IssueDescriptionList of Keyword
                        //    output: write color_description_list at Excel(row_index,new_inserted_col outside printable area
                        //         
                        // 3.3.1 Insert extra column just outside printable area.
                        // Assummed that Printable area always starting at $A$1 (also data processing area)
                        // So excel data processing area ends at Printable area (row_count,col_count)
                        int column_print_area = ExcelAction.GetWorksheetPrintableRange(result_worksheet).Columns.Count;
                        int insert_col = column_print_area + 1;
                        ExcelAction.Insert_Column(result_worksheet, insert_col);

                        // 3.3.2 Write keyword-related formatted issue descriptions on the newly-inserted column of the row where the keyword is found.
                        foreach (TestPlanKeyword keyword in keyword_list)
                        {
                            // Only write to keyword on currently open sheet
                            if (keyword.Worksheet == sheet_name)
                            {
                                int at_row = keyword.AtRow;
                                StyleString.WriteStyleString(result_worksheet, at_row, insert_col, keyword.IssueDescriptionList);
                            }
                        }

                        // 3.4. Save as another file with yyyyMMddHHmmss
                        string dest_filename = Storage.GenerateFilenameWithDateTime(full_filename);
                        ExcelAction.CloseExcelWorkbook(wb_keyword_issue, SaveChanges: true, AsFilename: dest_filename);
                    } 

                    return true;
                }
        */

        //static public bool KeywordIssueGenerationTaskV4(string report_filename)
        //{
        //    List<String> report_filename_list = new List<String>();
        //    report_filename_list.Add(report_filename);
        //    bool bRet = KeywordIssueGenerationTaskV4(report_filename_list, Storage.GetDirectoryName(report_filename));
        //    return bRet;
        //}

        static public void WriteBugCountOnKeywordReport(TestReportKeyword keyword, Worksheet result_worksheet, IssueCount severity_count)
        {
            // Write severity count of all keywrod isseus
            List<StyleString> bug_status_string = new List<StyleString>();
            int issue_count;
            issue_count = severity_count.Severity_A;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "A", Issue.KeywordIssue_A_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0A", Issue.KeywordIssue_report_FontColor));
            }
            //bug_status_string.Add(new StyleString(",", Color.Black));
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.Severity_B;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "B", Issue.KeywordIssue_B_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0B", Issue.KeywordIssue_report_FontColor));
            }
            //bug_status_string.Add(new StyleString(",", Color.Black));
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 1, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.Severity_C;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "C", Issue.KeywordIssue_C_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0C", Issue.KeywordIssue_report_FontColor));
            }
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 2, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.Severity_D;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "D", Issue.KeywordIssue_D_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0D", Issue.KeywordIssue_report_FontColor));
            }
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 3, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.TotalWaived();
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + " Waived", Issue.KeywordIssue_WAIVED_ISSUE_COLOR));
                StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 4, bug_status_string);
            }
            else
            {
                //bug_status_string.Add(new StyleString("No Waived", Color.Black));
                //StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 4, bug_status_string);
            }
            //StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 4, bug_status_string);
            bug_status_string.Clear();
        }

        static public void GetKeywordConclusionResult(IssueCount severity_count, out Boolean pass, out Boolean fail, out Boolean conditional_pass)
        {
            pass = fail = conditional_pass = false;

            if (severity_count.NotClosedCount() == 0)
            {
                // all issue closed
                pass = true;
            }
            else if (severity_count.ABC_non_Wavied_IssueCount() > 0)
            {
                // any issue of ABC, non-closed & non-waived issue 
                fail = true;
            }
            else
            {
                // only D or waived issue
                conditional_pass = true;
            }
        }

        static public void WriteKeywordConclusionOnKeywordReport(TestReportKeyword keyword, Worksheet result_worksheet, IssueCount severity_count)
        {
            String pass_fail_str;
            Boolean pass, fail, conditional_pass;
            GetKeywordConclusionResult(severity_count, out pass, out fail, out conditional_pass);

            if (pass == true)
            {
                pass_fail_str = PASS_str;
            }
            else if (fail == true)
            {
                pass_fail_str = FAIL_str;
            }
            else
            {
                pass_fail_str = CONDITIONAL_PASS_str;
            }
            ExcelAction.SetCellValue(result_worksheet, keyword.ResultAtRow, keyword.ResultAtColumn, pass_fail_str);
        }

        static public int FindTitle_Conclusion(Worksheet ws)
        {
            // find conclusion, if not found, row_default_conclusion_title as default
            int conclusion_start_row = row_test_brief_start, conclusion_end_row = row_test_brief_end;
            for (int row_index = conclusion_end_row; row_index >= conclusion_start_row; row_index--)
            {
                String text = ExcelAction.GetCellTrimmedString(ws, row_index, ExcelAction.ColumnNameToNumber("B"));
                if (CheckIfStringMeetsConclusion(text))
                {
                    return row_index;
                }
            }
            return row_default_conclusion_title;
        }

        static public Boolean ReplaceConclusionWithBugList(Worksheet ws, List<StyleString> bug_list_description)
        {
            int row_conclusion_title = FindTitle_Conclusion(ws);
            int row_bug_list_description = row_conclusion_title + 1;

            int col_start = ExcelAction.ColumnNameToNumber("C"),
                col_end = ExcelAction.ColumnNameToNumber("M");

            ExcelAction.ClearContent(ws, row_conclusion_title, col_start, row_bug_list_description, col_end);
            // output linked issue at C2
            StyleString.WriteStyleString(ws, row_conclusion_title + 1, 3, bug_list_description);
            ExcelAction.Merge(ws, row_conclusion_title + 1, col_start, row_bug_list_description, col_end);
            int line_count = 1; // at least one line, add one if "\n" encountered
            foreach (StyleString style_string in bug_list_description)
            {
                if (style_string.Text.Contains("\n"))
                {
                    line_count++;
                }
            }
            ExcelAction.Set_Row_Height(ws, row_bug_list_description, (StyleString.default_size + 1) * 2 * line_count * 0.75);
            return true;
        }

        static private Boolean CheckIfStringMeetsKeywordResultCondition(String text_to_check)
        {
            String regex = regexResultString;
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsKeywordBugStatusCondition(String text_to_check)
        {
            String regex = regexBugStatusString;
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsKeywordBugListCondition(String text_to_check)
        {
            String regex = regexBugListString;
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsTestPeriod(String text_to_check)
        {
            String regex = @"^(?i)\s*Test Period\s*$";
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsCriteria(String text_to_check)
        {
            String regex = @"^(?i)\s*Criteria:\s*$";
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsPurpose(String text_to_check)
        {
            String regex = @"^(?i)\s*Purpose:\s*$";
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsConclusion(String text_to_check)
        {
            String regex = @"^(?i)\s*Conclusion:\s*$";
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static public Boolean CheckIfStringMeetsMethod(String text_to_check)
        {
            String regex = @"^(?i)\s*Method:\s*$";
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsRegexString(String text_to_check, String regex_to_check)
        {
            Boolean ret_bol;
            RegexStringValidator RegexString = new RegexStringValidator(regex_to_check);
            try
            {
                RegexString.Validate(text_to_check);
                ret_bol = true;
            }
            catch (ArgumentException ex)
            {
                // does not meet
                ret_bol = false;
            }
            return ret_bol;
        }

        static private Boolean CheckIfStringMeetsSampleSN(String text_to_check)
        {
            String regex = @"^(?i)\s*Sample S/N:\s*$";
            return CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        // Code for Report *.0
        static public int Group_Summary_Table_RowNumber_Min = 3;
        public static int GroupSummary_Title_No_Row = 25, GroupSummary_Title_No_Col = ExcelAction.ColumnNameToNumber('D');
        public static int GroupSummary_Title_TestItem_Row = 25, GroupSummary_Title_TestItem_Col = ExcelAction.ColumnNameToNumber('E');
        public static int GroupSummary_Title_Result_Row = 25, GroupSummary_Title_Result_Col = ExcelAction.ColumnNameToNumber('H');
        public static int GroupSummary_Title_Note_Row = 25, GroupSummary_Title_Note_Col = ExcelAction.ColumnNameToNumber('J');
        public static String GroupSummary_Title_No_str = "No";
        public static String GroupSummary_Title_TestItem_str = "Test Item";
        public static String GroupSummary_Title_Result_str = "Result";
        public static String GroupSummary_Title_Note_str = "Note";

        static public Boolean Update_Single_Group_Summary_Report(Worksheet ws_group_report)
        {
            // check content of title row of group summary area, if not valid content, go to next
            if ((ExcelAction.CompareString(ws_group_report, GroupSummary_Title_No_Row, GroupSummary_Title_No_Col, GroupSummary_Title_No_str) == false) ||
                (ExcelAction.CompareString(ws_group_report, GroupSummary_Title_TestItem_Row, GroupSummary_Title_TestItem_Col, GroupSummary_Title_TestItem_str) == false) ||
                (ExcelAction.CompareString(ws_group_report, GroupSummary_Title_Result_Row, GroupSummary_Title_Result_Col, GroupSummary_Title_Result_str) == false) ||
                (ExcelAction.CompareString(ws_group_report, GroupSummary_Title_Note_Row, GroupSummary_Title_Note_Col, GroupSummary_Title_Note_str) == false))
            {
                return false;
            }

            return true;
        }


        static public bool Update_Group_Summary(Worksheet ws_report, String summary_report_sheetname, List<String> available_report_filelist)
        {
            Boolean b_ret = false;

            // find the table location

            if ((ExcelAction.CompareString(ws_report, GroupSummary_Title_No_Row, GroupSummary_Title_No_Col, GroupSummary_Title_No_str) == false) ||
                (ExcelAction.CompareString(ws_report, GroupSummary_Title_TestItem_Row, GroupSummary_Title_TestItem_Col, GroupSummary_Title_TestItem_str) == false) ||
                (ExcelAction.CompareString(ws_report, GroupSummary_Title_Result_Row, GroupSummary_Title_Result_Col, GroupSummary_Title_Result_str) == false) ||
                (ExcelAction.CompareString(ws_report, GroupSummary_Title_Note_Row, GroupSummary_Title_Note_Col, GroupSummary_Title_Note_str) == false))
            {
                return false;
            }

            // Find out all TCs required by this summary report
            String[] sp_str;
            sp_str = summary_report_sheetname.Split(new Char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            String target_prefix = sp_str[0];

            List<String> sub_report_list = new List<String>();
            foreach (String filename in available_report_filelist)
            {
                String filename_no_extension = Storage.GetFileNameWithoutExtension(filename);
                String report_sheetname = TestPlan.GetSheetNameAccordingToSummary(filename_no_extension);
                // same x.0, skip to next one
                if (report_sheetname == summary_report_sheetname)
                { continue; }

                sp_str = report_sheetname.Split(new Char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                String summary_prefix = sp_str[0];
                if (target_prefix != summary_prefix)
                { continue; }

                sub_report_list.Add(filename);
            }

            sub_report_list.Sort(TestPlan.Compare_Sheetname_by_Filename_Ascending);

            // Adjust and check the last row of table
            int row_found = 1;
            // target is the larger one between (1) number of rows required (2) Group_Summary_Table_RowNumber_Min
            int target_row_number = (sub_report_list.Count > Group_Summary_Table_RowNumber_Min) ?
                                    sub_report_list.Count : Group_Summary_Table_RowNumber_Min;

            do
            {
                int row_position = GroupSummary_Title_Note_Row + row_found;
                String no_str = ExcelAction.GetCellTrimmedString(ws_report, row_position, GroupSummary_Title_No_Col);
                if (String.IsNullOrWhiteSpace(no_str))
                {
                    break;
                }
                row_found++;
            }
            while (row_found <= target_row_number);

            // Adjust table row#
            // CASE 1: table is larger than target (Group_Summary_Table_RowNumber_Min or more) ==> reduce to target
            if (row_found > target_row_number)
            {
                // reduce row until target
                String no_str;
                int row_position = GroupSummary_Title_Note_Row + target_row_number + 1; // expected last row + 1 ==< expecting target+1 is empty
                do
                {
                    ExcelAction.Delete_Row(ws_report, GroupSummary_Title_Note_Row + (Group_Summary_Table_RowNumber_Min - 1));
                    no_str = ExcelAction.GetCellTrimmedString(ws_report, row_position, GroupSummary_Title_No_Col);
                }
                while (String.IsNullOrWhiteSpace(no_str) == false); // until last_row+1 is empty ==> last_row is the real last row
            }
            // CASE 2: table is smaller than target (3 or more) 
            //   2.1 target is smaller than or equal to Group_Summary_Table_RowNumber_Min, no need to increase
            //   2.2 increase rows until target
            else if (row_found < target_row_number)
            {
                // If 2.1 not met, executing 2.2
                if (target_row_number > Group_Summary_Table_RowNumber_Min)
                {
                    // need to increase row until target
                    do
                    {
                        ExcelAction.Insert_Row(ws_report, GroupSummary_Title_Note_Row + (Group_Summary_Table_RowNumber_Min - 1));
                        row_found++;
                    }
                    while (row_found < target_row_number);
                }
            }

            //update content 
            int row_index = GroupSummary_Title_Note_Row + 1;
            foreach (String filename in sub_report_list)
            {
                int col_index = GroupSummary_Title_No_Col;
                String str_no = TestPlan.GetSheetNameAccordingToFilename(filename);
                String str_test_item = TestPlan.GetReportTitleWithoutNumberAccordingToFilename(filename);
                String str_judgement = TestReport.Judgement_Decision_by_TC_Linked_Issue(str_no);
                List<StyleString> linked_issue_description = new List<StyleString>();
                if (ReportGenerator.GetTestcaseLUT_by_Sheetname().ContainsKey(str_no))
                {
                    // key string of all linked issue
                    String links = ReportGenerator.GetTestcaseLUT_by_Sheetname()[str_no].Links;
                    // key string to List of Issue
                    List<Issue> linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                    // List of Issue filtered by status
                    List<Issue> filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.List_of_status_to_filter_for_tc_linked_issue);
                    // Sort issue by Severity and Key valie
                    List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
                    // Convert list of sorted linked issue to description list
                    linked_issue_description = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);
                }
                ExcelAction.SetCellValue(ws_report, row_index, GroupSummary_Title_No_Col, str_no);
                ExcelAction.SetCellValue(ws_report, row_index, GroupSummary_Title_TestItem_Col, str_test_item);
                ExcelAction.SetCellValue(ws_report, row_index, GroupSummary_Title_Result_Col, str_judgement, ClearContentFirst: true);
                StyleString.WriteStyleString(ws_report, row_index, GroupSummary_Title_Note_Col, linked_issue_description, ClearContentFirst: true);
                row_index++;
            }
            target_row_number += GroupSummary_Title_Note_Row;
            while (row_index <= target_row_number)
            {
                int col_index = GroupSummary_Title_No_Col;
                ExcelAction.SetCellValue(ws_report, row_index, GroupSummary_Title_No_Col, " ");
                ExcelAction.SetCellValue(ws_report, row_index, GroupSummary_Title_TestItem_Col, " ");
                ExcelAction.SetCellValue(ws_report, row_index, GroupSummary_Title_Result_Col, " ", ClearContentFirst: true);
                StyleString.WriteStyleString(ws_report, row_index, GroupSummary_Title_Note_Col, StyleString.EmptyList(), ClearContentFirst: true);
                row_index++;
            }

            b_ret = true;

            return b_ret;
        }

        static public bool Update_Group_Summary(String report_path)
        {
            Boolean b_ret = false;
            String destination_path = Storage.GenerateDirectoryNameWithDateTime(report_path);

            // 1. List excel under report_path
            // 2. keep only "x.0" on the file list
            List<String> group_file_list = Storage.ListCandidateGroupSummaryFilesUnderDirectory(report_path);

            foreach (String group_file in group_file_list)
            {
                // 3. open
                // open standard test report
                Workbook wb_report = ExcelAction.OpenExcelWorkbook(group_file);
                if (wb_report == null)
                {
                    LogMessage.WriteLine("OpenExcelWorkbook failed in Update_Group_Summary()");
                    continue;
                }

                String sheet_name = TestPlan.GetSheetNameAccordingToFilename(group_file);
                // Select and read work-sheet
                Worksheet ws_report = ExcelAction.Find_Worksheet(wb_report, sheet_name);
                if (ws_report == null)
                {
                    LogMessage.WriteLine("Find_Worksheet (" + sheet_name + ") failed in Update_Group_Summary()");
                    ExcelAction.CloseExcelWorkbook(wb_report);
                    continue;
                }

                // 4. check content of title row, if not valid content, go to next
                if ((ExcelAction.CompareString(ws_report, GroupSummary_Title_No_Row, GroupSummary_Title_No_Col, GroupSummary_Title_No_str) == false) ||
                    (ExcelAction.CompareString(ws_report, GroupSummary_Title_TestItem_Row, GroupSummary_Title_TestItem_Col, GroupSummary_Title_TestItem_str) == false) ||
                    (ExcelAction.CompareString(ws_report, GroupSummary_Title_Result_Row, GroupSummary_Title_Result_Col, GroupSummary_Title_Result_str) == false) ||
                    (ExcelAction.CompareString(ws_report, GroupSummary_Title_Note_Row, GroupSummary_Title_Note_Col, GroupSummary_Title_Note_str) == false))
                {
                    ExcelAction.CloseExcelWorkbook(wb_report);
                    continue;
                }

                // 6. adjust table rows according to the number of TC summary within this gruop
                // count tc case in this group
                int tc_count = 0;
                if (tc_count == 0)
                {
                    ExcelAction.CloseExcelWorkbook(wb_report);
                    continue;
                }

                // find table lower boundary
                int row_index = GroupSummary_Title_No_Row + 1;
                int row_end = ExcelAction.Get_Range_RowNumber(ExcelAction.GetWorksheetAllRange(ws_report));
                Boolean row_found = false;
                while (!row_found && (row_index <= row_end))
                {
                    if (ExcelAction.GetCellTrimmedString(ws_report, row_index, GroupSummary_Title_No_Col) == "")
                    {
                        row_found = true;
                        row_index--;
                    }
                    else
                    {
                        row_index++;
                    }
                }

                // adjust row number


                // 7. Fill each row with (1) sheetname (2) summary after sheetname (3) TC result (4) linked issue on TC
                // for each tc item
                int current_row = 111;

                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_No_Col, "NO");
                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_TestItem_Col, "name");
                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_Result_Col, "judgement");
                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_Note_Col, "linked issue");

                // 8. save and close and go to next report
                String output_filename = group_file.Replace(report_path, destination_path);
                ExcelAction.SaveExcelWorkbook(wb_report, output_filename);
                ExcelAction.CloseExcelWorkbook(wb_report);
            }

            b_ret = true;
            return b_ret;
        }

        // Extracted from KeywordIssueGenerationTaskV4()
        static public bool KeywordIssueGenerationTaskV4_simplified_SingleReport_Processing(Worksheet result_worksheet, List<String> existing_report_filelist, String dest_filename)
        {
            String sheet_name = result_worksheet.Name;
            // if sheetname is xxxxxxx.0, do group_summary_report)
            if (sheet_name.Substring(sheet_name.Length - 2, 2) == ".0")
            {
                TestReport.Update_Group_Summary(result_worksheet, sheet_name, existing_report_filelist);
            }

            String judgement_str = "", purpose_str = "", criteria_str = "";
            // 3.3_minus 1: store text of purpose & criteia for updating into TC summary report
            // always stored even it is not a key-word report.
            int search_start_row = row_test_brief_start, search_end_row = row_test_brief_end;
            int purpose_str_col = ExcelAction.ColumnNameToNumber('C');
            int criteria_str_col = ExcelAction.ColumnNameToNumber('C');
            Boolean purpose_found = false, criteria_found = false;
            for (int row_index = search_start_row; row_index <= search_end_row; row_index++)
            {
                String text = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, col_indentifier);
                if (purpose_found == false)
                {
                    if (CheckIfStringMeetsPurpose(text))
                    {
                        row_index++;
                        purpose_str = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, purpose_str_col);
                        purpose_found = true;
                        if (criteria_found)
                            break;
                        else
                            continue;
                    }
                }
                if (criteria_found == false)
                {
                    if (CheckIfStringMeetsCriteria(text))
                    {
                        row_index++;
                        criteria_str = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, criteria_str_col);
                        criteria_found = true;
                        if (purpose_found)
                            break;
                        else
                            continue;
                    }
                }
            }
            judgement_str = ExcelAction.GetCellTrimmedString(result_worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col);

            //if (Replace_Conclusion)
            if (true)    // always updating linked issue for non-keyword version of report 
            {
                Update_Conclusion_Judgement_by_linked_issue(result_worksheet);

                /*
                // Add: replace conclusion with Bug-list
                //ReplaceConclusionWithBugList(result_worksheet, keyword_issue_description_on_this_report); // should be linked issue in the future
                // Find the TC meets the sheet-name
                List<StyleString> linked_issue_description_on_this_report = new List<StyleString>();
                if (ReportGenerator.GetTestcaseLUT_by_Sheetname().ContainsKey(sheet_name))          // if TC is available
                {
                    // key string of all linked issues
                    String links = ReportGenerator.GetTestcaseLUT_by_Sheetname()[sheet_name].Links;
                    // key string to List of Issue
                    List<Issue> linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                    // List of Issue filtered by status
                    List<Issue> filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.filter_status_list_linked_issue);
                    // Sort issue by Severity and Key valie
                    List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
                    // Convert list of sorted linked issue to description list
                    linked_issue_description_on_this_report = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);

                    // decide judgement result based on linked issue severity and count
                    judgement_str = Judgement_Decision_by_Linked_Issue(linked_issue_list);
                }
                else
                {
                    linked_issue_description_on_this_report.Clear();
                    linked_issue_description_on_this_report.Add(KeywordReportHeader.blank_space);
                }
                ReplaceConclusionWithBugList(result_worksheet, linked_issue_description_on_this_report);
                //
                ExcelAction.CellActivate(result_worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col);
                ExcelAction.SetCellValue(result_worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col, judgement_str);
                */
            }
            // always update Test End Period to today
            if (true)      // this part of code is only for old header mechanism before header template is available
            {
                String end_date = DateTime.Now.ToString("yyyy/MM/dd");
                String text_to_check;
                int today_row = 0, today_col = 0, check_row, check_col;
                // Check format 1
                check_row = 8;
                check_col = ExcelAction.ColumnNameToNumber('J');
                text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, check_row, check_col);
                if (CheckIfStringMeetsTestPeriod(text_to_check))
                {
                    today_row = 8;
                    today_col = ExcelAction.ColumnNameToNumber('M');
                }
                else
                {
                    // Check format 1
                    check_row = 8;
                    check_col = ExcelAction.ColumnNameToNumber('H');
                    text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, check_row, check_col);
                    if (CheckIfStringMeetsTestPeriod(text_to_check))
                    {
                        today_row = 8;
                        today_col = ExcelAction.ColumnNameToNumber('L');
                        end_date = "-             " + end_date;
                    }
                }
                if ((today_row > 0) && (today_col > 0))
                {
                    ExcelAction.SetCellValue(result_worksheet, today_row, today_col, end_date);
                }
            }
            //// update Part No.
            //String default_part_no = "99.M2710.0A4-";
            //String part_no = default_part_no + sheet_name;
            //ExcelAction.SetCellValue(result_worksheet, Part_No_at_row, Part_No_at_col, part_no);

            List<String> report_info = CombineReportInfo(judgement: judgement_str, purpose: purpose_str, criteria: criteria_str);
            AppendReportInformation(dest_filename, report_info);

            return true;
        }

        static public bool KeywordIssueGenerationTaskV4_simplified(List<String> file_list, String src_dir, String dest_dir = "")
        {
            // 0.1 List all files under report_root_dir.
            // This is done outside and result is the input paramemter file_list
            // 0.2 filename check to exclude non-report files.
            List<String> existing_report_filelist = Storage.FilterFilename(file_list);
            // Sorting report_file (file_list) in descending order so that x.0 report will be processed after all other x.n reprot
            existing_report_filelist.Sort(TestPlan.Compare_Sheetname_by_Filename_Descending);

            //
            // 1. Create a temporary test plan (do_plan) to include all report files 
            //
            // 1.1 Init an empty plan
            List<TestPlan> do_plan = new List<TestPlan>();
            ClearReportInformation();

            // 1.2 Create a temporary test plan to includes all files listed in List<String> report_filename
            do_plan = TestPlan.CreateTempPlanFromFileList(existing_report_filelist);

            foreach (TestPlan plan in do_plan)
            {
                String full_filename = plan.ExcelFile;
                String sheet_name = plan.ExcelSheet;
                // 3.1. Open Excel and find the sheet
                // File exist check is done outside
                Workbook wb_keyword_issue = ExcelAction.OpenExcelWorkbook(full_filename);
                if (wb_keyword_issue == null)
                {
                    LogMessage.WriteLine("ERR: Open workbook in V4: " + full_filename);
                    continue;
                }

                // 3.2 Open worksheet
                Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_keyword_issue, sheet_name);
                if (result_worksheet == null)
                {
                    LogMessage.WriteLine("ERR: Open worksheet in V4: " + full_filename + " sheet: " + sheet_name);
                    continue;
                }

                String dest_filename = DecideDestinationFilename(src_dir, dest_dir, full_filename);
                KeywordIssueGenerationTaskV4_simplified_SingleReport_Processing(result_worksheet, existing_report_filelist, dest_filename);

                // 3.4. Save the file to either 
                //  (1) filename with yyyyMMddHHmmss if dest_dir is not specified
                //  (2) the same filename but to the sub-folder of same strucure under new root-folder "dest_dir"
                String dest_filename_dir = Storage.GetDirectoryName(dest_filename);
                // if parent directory does not exist, create recursively all parents
                Storage.CreateDirectory(dest_filename_dir, auto_parent_dir: true);
                ExcelAction.CloseExcelWorkbook(wb_keyword_issue, SaveChanges: true, AsFilename: dest_filename);
            }
            return true;
        }

        static public bool KeywordIssueGenerationTaskV4(List<String> file_list, String src_dir, String dest_dir = "")
        {
            // Clear keyword log report data-table
            ReportGenerator.excel_not_report_log.Clear();
            // 0.1 List all files under report_root_dir.
            // This is done outside and result is the input paramemter file_list
            // 0.2 filename check to exclude non-report files.
            List<String> existing_report_filelist = Storage.FilterFilename(file_list);
            // Sorting report_file (file_list) in descending order so that x.0 report will be processed after all other x.n reprot
            existing_report_filelist.Sort(TestPlan.Compare_Sheetname_by_Filename_Descending);
            // 0.3 output files in file_list but not in report_filename into Not_Keyword_File
            foreach (String report_file in existing_report_filelist)
            {
                file_list.Remove(report_file);
            }
            foreach (String NG_file in file_list)
            {
                String path, filename;
                path = Storage.GetDirectoryName(NG_file);
                filename = Storage.GetFileName(NG_file);
                ReportFileRecord nrfr_item = new ReportFileRecord(path, filename);
                nrfr_item.SetFlagFail(excelfilenamefail: true);
                ReportGenerator.excel_not_report_log.Add(nrfr_item);
            }

            //
            // 1. Create a temporary test plan (do_plan) to include all report files 
            //
            // 1.1 Init an empty plan
            List<TestPlan> do_plan = new List<TestPlan>();

            // 1.2 Create a temporary test plan to includes all files listed in List<String> report_filename
            do_plan = TestPlan.CreateTempPlanFromFileList(existing_report_filelist);
            //ClearReportJudgementResult();
            ClearReportInformation();

            //
            // 2. Search keywords within all selected file (2.1) and use those keywords to find out issues containing keywords.
            //
            // 2.1. Find keyword for all selected file (as listed in temprary test plan)
            //
            List<TestReportKeyword> keyword_list = ListAllKeyword(do_plan);
            // Clear global_keyword_list here
            ClearGlobalKeywordList();

            //// Output keyword list log excel here.
            //String out_dir = (dest_dir!="")?dest_dir:src_dir;
            //KeyWordListReport.OutputKeywordLog(out_dir, keyword_list, ReportGenerator.excel_not_report_log);

            //
            // 2.2. Use keyword to find out all issues (ID) that contains keyword on id_list. 
            //    Extend list of issue ID to list of issue description (with font style settings) -- by Issue.GenerateIssueDescription
            //
            //ReportGenerator.global_issue_description_list = Issue.GenerateIssueDescription(ReportGenerator.global_issue_list);//done outside in advance
            Dictionary<String, List<String>> KeywordIssueIDList = new Dictionary<String, List<String>>();
            foreach (Issue issue in ReportGenerator.ReadGlobalIssueList())
            {
                issue.KeywordList.Clear();
            }
            // Go throught each keyword, search all issues containing this keyword and add issue-id so that it can be extened into description list.
            foreach (TestReportKeyword keyword in keyword_list)
            {
                List<StyleString> description_list;
                //List<String> id_list = new List<String>();
                keyword.KeywordIssues.Clear();
                String keyword_str = keyword.Keyword;
                foreach (Issue issue in ReportGenerator.ReadGlobalIssueList())
                {
                    // if status meets filter condition (mostly Closed_0), skip to next issue)
                    if (KeywordIssue_filter_status_list.IndexOf(issue.Status) >= 0)
                    {
                        continue;
                    }
                    if (issue.ContainKeyword(keyword_str))
                    {
                        issue.KeywordList.Add(keyword_str);
                        keyword.KeywordIssues.Add(issue);       // keep issue with keyword so that it can be used later.
                    }
                }
                //keyword.IssueList = Issue.ListOfIssueToListOfIssueKey(keyword.KeywordIssues);
                // Sort issue by Severity and Key valie
                List<Issue> sorted_keyword_issues = Issue.SortingBySeverityAndKey(keyword.KeywordIssues);
                description_list = StyleString.BugList_To_KeywordIssueDescription(sorted_keyword_issues);
                keyword.IssueDescriptionList = description_list;
            }

            // Output keyword list log excel here.
            String out_dir = (String.IsNullOrWhiteSpace(dest_dir) == false) ? dest_dir : src_dir;
            KeyWordListReport.OutputKeywordLog(out_dir, keyword_list, ReportGenerator.excel_not_report_log, output_keyword_issue: true);
            //KeyWordListReport.OutputKeywordLog(out_dir, keyword_list, ReportGenerator.excel_not_report_log, output_keyword_issue: false);

            // Load global_keyword_list here
            //global_keyword_list = keyword_list;
            //global_keyword_available = true;
            SetGlobalKeywordList(keyword_list);
            Dictionary<String, List<TestReportKeyword>> keyword_lut_by_sheetname = GenerateKeywordLUT_by_Sheetname(keyword_list);

            //
            // 3. Go throught each report excel and generate keyword report for each one.
            //
            foreach (TestPlan plan in do_plan)
            {
                String full_filename = plan.ExcelFile;
                String sheet_name = plan.ExcelSheet;

                // 3.1. Open Excel and find the sheet
                // File exist check is done outside
                Workbook wb_keyword_issue = ExcelAction.OpenExcelWorkbook(full_filename);
                if (wb_keyword_issue == null)
                {
                    LogMessage.WriteLine("ERR: Open workbook in V4: " + full_filename);
                    continue;
                }

                // 3.2 Open worksheet
                Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_keyword_issue, sheet_name);
                if (result_worksheet == null)
                {
                    LogMessage.WriteLine("ERR: Open worksheet in V4: " + full_filename + " sheet: " + sheet_name);
                    continue;
                }

                // if sheetname is xxxxxxx.0, do group_summary_report)
                if (sheet_name.Substring(sheet_name.Length - 2, 2) == ".0")
                {
                    TestReport.Update_Group_Summary(result_worksheet, sheet_name, existing_report_filelist);
                }

                String judgement_str = "", purpose_str = "", criteria_str = "";
                // 3.3_minus 1: store text of purpose & criteia for updating into TC summary report
                // always stored even it is not a key-word report.
                int search_start_row = row_test_brief_start, search_end_row = row_test_brief_end;
                int purpose_str_col = ExcelAction.ColumnNameToNumber('C');
                int criteria_str_col = ExcelAction.ColumnNameToNumber('C');
                Boolean purpose_found = false, criteria_found = false;
                for (int row_index = search_start_row; row_index <= search_end_row; row_index++)
                {
                    String text = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, col_indentifier);
                    if (purpose_found == false)
                    {
                        if (CheckIfStringMeetsPurpose(text))
                        {
                            row_index++;
                            purpose_str = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, purpose_str_col);
                            purpose_found = true;
                            if (criteria_found)
                                break;
                            else
                                continue;
                        }
                    }
                    if (criteria_found == false)
                    {
                        if (CheckIfStringMeetsCriteria(text))
                        {
                            row_index++;
                            criteria_str = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, criteria_str_col);
                            criteria_found = true;
                            if (purpose_found)
                                break;
                            else
                                continue;
                        }
                    }
                }
                judgement_str = ExcelAction.GetCellTrimmedString(result_worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col);

                if (keyword_lut_by_sheetname.ContainsKey(sheet_name) == true)
                {
                    // if keyword exist, executing keyword-related part
                    // 
                    // 3.3. input:  IssueDescriptionList of Keyword
                    //    output: write color_description_list 
                    //         

                    //// 3.3.0: store text of purpose & criteia for updating into TC summary report
                    //int search_start_row = row_test_brief_start, search_end_row = row_test_brief_end;
                    //for (int row_index = search_start_row; row_index <= search_end_row; row_index++)
                    //{
                    //    String text = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, col_indentifier);
                    //    if(CheckIfStringMeetsPurpose(text))
                    //    {
                    //        row_index++;
                    //        purpose_str = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, ExcelAction.ColumnNameToNumber('C'));
                    //        continue;
                    //    }
                    //    else if (CheckIfStringMeetsCriteria(text))
                    //    {
                    //        row_index++;
                    //        criteria_str = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, ExcelAction.ColumnNameToNumber('C'));
                    //        continue;
                    //    }
                    //}

                    // 3.3.2 Write keyword-related formatted issue descriptions 
                    //       also count how many "Pass" or how many "Fail"


                    int pass_count = 0, fail_count = 0, conditional_pass_count = 0;
                    //foreach (TestPlanKeyword keyword in keyword_list)
                    //foreach (TestPlanKeyword keyword in ws_keyword_list)
                    List<StyleString> keyword_issue_description_on_this_report = new List<StyleString>();
                    foreach (TestReportKeyword keyword in keyword_lut_by_sheetname[sheet_name])
                    {
                        // Only write to keyword on currently open sheet
                        //if (keyword.Worksheet == sheet_name)
                        {
                            // write issue description list

                            // Because keyword condition is now relaxed and "Bug-List"..etc is no longer part of condition,
                            // It needs to be checked here before content is written
                            // Check BugList
                            String text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn - 1);
                            if (CheckIfStringMeetsKeywordBugListCondition(text_to_check))
                            {
                                StyleString.WriteStyleString(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn, keyword.IssueDescriptionList);
                                ExcelAction.AutoFit_Row(result_worksheet, keyword.BugListAtRow);
                                keyword_issue_description_on_this_report.AddRange(keyword.IssueDescriptionList);
                            }

                            // write issue count of each severity
                            // IssueCount severity_count = keyword.Calculate_Issue(); 
                            IssueCount severity_count = IssueCount.IssueListStatistic(keyword.KeywordIssues);
                            // Check bug_status
                            text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn - 1);
                            if (CheckIfStringMeetsKeywordBugStatusCondition(text_to_check))
                            {
                                WriteBugCountOnKeywordReport(keyword, result_worksheet, severity_count);
                                ExcelAction.AutoFit_Row(result_worksheet, keyword.BugStatusAtRow);
                            }

                            // write conclusion of each keyword
                            Boolean pass, fail, conditional_pass;
                            GetKeywordConclusionResult(severity_count, out pass, out fail, out conditional_pass);
                            // Check BugStatus
                            text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, keyword.ResultAtRow, keyword.ResultAtColumn - 1);
                            if (CheckIfStringMeetsKeywordResultCondition(text_to_check))
                            {
                                WriteKeywordConclusionOnKeywordReport(keyword, result_worksheet, severity_count);
                                ExcelAction.AutoFit_Row(result_worksheet, keyword.ResultAtRow);
                            }

                            if (pass)
                            {
                                pass_count++;
                            }
                            else if (fail)
                            {
                                fail_count++;
                            }
                            else
                            {
                                conditional_pass_count++;
                            }

                            // auto-fit row-height
                            // Check cell, auto-fit if met
                            //ExcelAction.AutoFit_Row(result_worksheet, keyword.ResultAtRow);
                            //ExcelAction.AutoFit_Row(result_worksheet, keyword.BugListAtRow);
                            // issue_count = severity_count.Severity_A + severity_count.Severity_B + severity_count.Severity_C;
                            //if (issue_count >= 1)
                            text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn - 1);
                            if (CheckIfStringMeetsKeywordBugListCondition(text_to_check))
                            {
                                int issue_count = severity_count.NotClosedCount();
                                if (issue_count > 0)
                                {
                                    double single_row_height = (StyleString.default_size + 1) * 2 * 0.75;
                                    double new_row_height = single_row_height * issue_count;
                                    // Check cell, unhide if met
                                    ExcelAction.Set_Row_Height(result_worksheet, keyword.BugListAtRow, new_row_height);
                                }
                                else
                                {
                                    // Hide bug list row only when there isn't any non-closed issue at all (all issues must be closed)
                                    double new_row_height = 0.2;
                                    // Check cell, hide if met
                                    ExcelAction.Set_Row_Height(result_worksheet, keyword.BugListAtRow, new_row_height);
                                    //ExcelAction.Hide_Row(result_worksheet, keyword.BugListAtRow);
                                }
                                //ExcelAction.CellTextAlignLeft(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn);
                                ExcelAction.CellTextAlignUpperLeft(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn);
                            }

                            if (KeywordIssue_Hide_Result_Bug)
                            {
                                double new_row_height = 0.2;
                                // Check cell, hide if met
                                text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn - 1);
                                if (CheckIfStringMeetsKeywordBugListCondition(text_to_check))
                                {
                                    ExcelAction.Set_Row_Height(result_worksheet, keyword.BugListAtRow, new_row_height);
                                }

                                // Check cell, hide if met
                                text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, keyword.ResultAtRow, keyword.ResultAtColumn - 1);
                                if (CheckIfStringMeetsKeywordResultCondition(text_to_check))
                                {
                                    ExcelAction.Set_Row_Height(result_worksheet, keyword.ResultAtRow, new_row_height);
                                }

                                // Check cell, hide if met
                                text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn - 1);
                                if (CheckIfStringMeetsKeywordBugStatusCondition(text_to_check))
                                {
                                    ExcelAction.Set_Row_Height(result_worksheet, keyword.BugStatusAtRow, new_row_height);
                                }
                            }
                        }
                    }

                    // 3.3.3 Update Conclusion
                    //const int PassCnt_at_row = 21, PassCnt_at_col = 5;
                    //const int FailCnt_at_row = 21, FailCnt_at_col = 7;
                    //const int TotalCnt_at_row = 21, TotalCnt_at_col = 9;
                    //const int Judgement_at_row = 9, Judgement_at_col = 4;
                    if (fail_count > 0)
                    {
                        // Fail
                        judgement_str = FAIL_str;
                    }
                    else if (conditional_pass_count > 0)
                    {
                        // conditional pass
                        judgement_str = CONDITIONAL_PASS_str;
                    }
                    else
                    {
                        // pass
                        judgement_str = PASS_str;
                    }

                    ExcelAction.SetCellValue(result_worksheet, PassCnt_at_row, PassCnt_at_col, pass_count);
                    ExcelAction.SetCellValue(result_worksheet, FailCnt_at_row, FailCnt_at_col, fail_count);
                    ExcelAction.SetCellValue(result_worksheet, ConditionalPass_string_at_row, ConditionalPass_string_at_col, CONDITIONAL_PASS_str + ":");
                    ExcelAction.SetCellValue(result_worksheet, ConditionalPassCnt_at_row, ConditionalPassCnt_at_col, conditional_pass_count);
                    ExcelAction.SetCellValue(result_worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col, judgement_str);

                    // End of updating keyword result
                }

                ////Report_C_Update_Header_by_Template
                //if (KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template == true)
                //{
                //    HeaderTemplate.ReplaceHeaderVariableWithValue(result_worksheet);
                //}

                if (KeywordIssue_Replace_Conclusion)
                {
                    Update_Conclusion_Judgement_by_linked_issue(result_worksheet);
                    /*
                    // Add: replace conclusion with Bug-list
                    //ReplaceConclusionWithBugList(result_worksheet, keyword_issue_description_on_this_report); // should be linked issue in the future
                    // Find the TC meets the sheet-name
                    List<StyleString> linked_issue_description_on_this_report = new List<StyleString>();
                    if (ReportGenerator.GetTestcaseLUT_by_Sheetname().ContainsKey(sheet_name))
                    {
                        // key string of all linked issues
                        String links = ReportGenerator.GetTestcaseLUT_by_Sheetname()[sheet_name].Links;
                        // key string to List of Issue
                        List<Issue> linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                        // List of Issue filtered by status
                        List<Issue> filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.filter_status_list_linked_issue);
                        // Sort issue by Severity and Key valie
                        List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
                        // Convert list of sorted linked issue to description list
                        linked_issue_description_on_this_report = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);

                        judgement_str = Judgement_Decision_by_Linked_Issue(linked_issue_list);
                    }
                    else
                    {
                        linked_issue_description_on_this_report.Clear();
                        linked_issue_description_on_this_report.Add(KeywordReportHeader.blank_space);
                    }
                    ReplaceConclusionWithBugList(result_worksheet, linked_issue_description_on_this_report);
                    //
                    ExcelAction.CellActivate(result_worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col);
                    ExcelAction.SetCellValue(result_worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col, judgement_str);
                    */ 
                }

                // always update Test End Period to today
                if (true)      // this part of code is only for old header mechanism before header template is available
                {
                    String end_date = DateTime.Now.ToString("yyyy/MM/dd");
                    String text_to_check;
                    int today_row = 0, today_col = 0, check_row, check_col;
                    // Check format 1
                    check_row = 8;
                    check_col = ExcelAction.ColumnNameToNumber('J');
                    text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, check_row, check_col);
                    if (CheckIfStringMeetsTestPeriod(text_to_check))
                    {
                        today_row = 8;
                        today_col = ExcelAction.ColumnNameToNumber('M');
                    }
                    else
                    {
                        // Check format 1
                        check_row = 8;
                        check_col = ExcelAction.ColumnNameToNumber('H');
                        text_to_check = ExcelAction.GetCellTrimmedString(result_worksheet, check_row, check_col);
                        if (CheckIfStringMeetsTestPeriod(text_to_check))
                        {
                            today_row = 8;
                            today_col = ExcelAction.ColumnNameToNumber('L');
                            end_date = "-             " + end_date;
                        }
                    }
                    if ((today_row > 0) && (today_col > 0))
                    {
                        ExcelAction.SetCellValue(result_worksheet, today_row, today_col, end_date);
                    }
                }
                //// update Part No.
                //String default_part_no = "99.M2710.0A4-";
                //String part_no = default_part_no + sheet_name;
                //ExcelAction.SetCellValue(result_worksheet, Part_No_at_row, Part_No_at_col, part_no);

                // 3.4. Save the file to either 
                //  (1) filename with yyyyMMddHHmmss if dest_dir is not specified
                //  (2) the same filename but to the sub-folder of same strucure under new root-folder "dest_dir"
                String dest_filename = DecideDestinationFilename(src_dir, dest_dir, full_filename);
                String dest_filename_dir = Storage.GetDirectoryName(dest_filename);
                // if parent directory does not exist, create recursively all parents
                Storage.CreateDirectory(dest_filename_dir, auto_parent_dir: true);
                ExcelAction.CloseExcelWorkbook(wb_keyword_issue, SaveChanges: true, AsFilename: dest_filename);

                //AppendReportJudgementResult(dest_filename, judgement_str);
                List<String> report_info = CombineReportInfo(judgement: judgement_str, purpose: purpose_str, criteria: criteria_str);
                AppendReportInformation(dest_filename, report_info);
            }

            // Output updated report with recommended sheetname.
            if (TestReport.KeywordIssue_Auto_Correct_Sheetname == true)
            {
                // ReportGenerator.excel_not_report_log
                foreach (ReportFileRecord item in ReportGenerator.excel_not_report_log)
                {
                    String path, filename, expected_sheetname;
                    Boolean excelfilenameOK, openfileOK, findWorksheetOK, findAnyKeyword, otherFailure;

                    item.GetRecord(out path, out filename, out expected_sheetname, out excelfilenameOK, out openfileOK, out findWorksheetOK,
                            out findAnyKeyword, out otherFailure);

                    if ((excelfilenameOK == true) && (openfileOK == true) && (findWorksheetOK == false) && (otherFailure == false))
                    {
                        String full_filename = Storage.GetValidFullFilename(path, filename);
                        // Open Excel and find the sheet
                        Workbook wb = ExcelAction.OpenExcelWorkbook(filename: full_filename, ReadOnly: false);
                        if (wb == null)
                        {
                            LogMessage.WriteLine("ERR: Open workbook in auto-correct-worksheet-name: " + full_filename);
                            continue;
                        }

                        // Use first worksheet and rename it.
                        Worksheet ws = wb.Sheets[1];
                        ws.Name = expected_sheetname;

                        // Save the updated report file file to either 
                        //  (1) filename with yyyyMMddHHmmss if dest_dir is not specified
                        //  (2) the same filename but to the sub-folder of same strucure under new root-folder "dest_dir"
                        String dest_filename = DecideDestinationFilename(src_dir, dest_dir, full_filename);
                        String dest_filename_dir = Storage.GetDirectoryName(dest_filename);
                        // if parent directory does not exist, create recursively all parents
                        Storage.CreateDirectory(dest_filename_dir, auto_parent_dir: true);
                        ExcelAction.CloseExcelWorkbook(wb, SaveChanges: true, AsFilename: dest_filename);
                    }
                }
            }

            return true;
        }

        // Please input linked issue
        static public String Judgement_Decision_by_Linked_Issue_List(List<Issue> linked_issue_list)
        {
            String judgement_str = TestReport.TestReport_Default_Judgement;

            // List of Issue filtered by status
            List<Issue> filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.List_of_status_to_filter_for_tc_linked_issue);
            // count of filtered issue
            IssueCount severity_count = IssueCount.IssueListStatistic(filtered_linked_issue_list);
            Boolean pass, fail, conditional_pass;
            GetKeywordConclusionResult(severity_count, out pass, out fail, out conditional_pass);
            if (fail)
            {
                judgement_str = FAIL_str;
            }
            else if (conditional_pass)
            {
                judgement_str = CONDITIONAL_PASS_str;
            }
            else
            {
                judgement_str = PASS_str;
            }
            return judgement_str;
        }

        static public String Judgement_Decision_by_TC_Linked_Issue(String sheet_name)
        {
            // return default if no corresponding TC
            String judgement_str = TestReport.TestReport_Default_Judgement;
            if (ReportGenerator.GetTestcaseLUT_by_Sheetname().ContainsKey(sheet_name))
            {
                // key string of all linked issue
                String links = ReportGenerator.GetTestcaseLUT_by_Sheetname()[sheet_name].Links;
                // key string to List of Issue
                List<Issue> linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                judgement_str = Judgement_Decision_by_Linked_Issue_List(linked_issue_list);
            }
            return judgement_str;
        }

        // to be finished
        static public Boolean ReadKeywordReportHeader_full(Worksheet report_worksheet, out TestReportOption out_header)
        {
            out_header = new TestReportOption();
            out_header.Report_Title = ExcelAction.GetCellTrimmedString(report_worksheet, TestReport.Title_at_row, TestReport.Title_at_col);
            out_header.Model_Name = ExcelAction.GetCellTrimmedString(report_worksheet, TestReport.Model_Name_at_row, TestReport.Model_Name_at_col);
            ////@"Update_Part_No",                          @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Part_No)
            //{
            //    String output_part_no = header.Part_No + "-" + header.Report_SheetName;
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Part_No_at_row, KeywordReportHeader.Part_No_at_col, output_part_no);
            //}

            ////@"Update_Panel_Module",                     @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Panel_Module)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Panel_Module_at_row, KeywordReportHeader.Panel_Module_at_col, header.Panel_Module);
            //}

            ////@"Update_TCON_Board",                       @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_TCON_Board)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.TCON_Board_at_row, KeywordReportHeader.TCON_Board_at_col, header.TCON_Board);
            //}

            ////@"Update_AD_Board",                         @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_AD_Board)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.AD_Board_at_row, KeywordReportHeader.AD_Board_at_col, header.AD_Board);
            //}

            ////@"Update_Power_Board",                      @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Power_Board)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Power_Board_at_row, KeywordReportHeader.Power_Board_at_col, header.Power_Board);
            //}

            ////@"Update_Smart_BD_OS_Version",              @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Smart_BD_OS_Version)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Smart_BD_OS_Version_at_row, KeywordReportHeader.Smart_BD_OS_Version_at_col, header.Smart_BD_OS_Version);
            //}

            ////@"Update_Touch_Sensor",                     @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Touch_Sensor)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Touch_Sensor_at_row, KeywordReportHeader.Touch_Sensor_at_col, header.Touch_Sensor);
            //}

            ////@"Update_Speaker_AQ_Version",               @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Speaker_AQ_Version)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Speaker_AQ_Version_at_row, KeywordReportHeader.Speaker_AQ_Version_at_col, header.Speaker_AQ_Version);
            //}

            ////@"Update_SW_PQ_Version",                    @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_SW_PQ_Version)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.SW_PQ_Version_at_row, KeywordReportHeader.SW_PQ_Version_at_col, header.SW_PQ_Version);
            //}

            ////@"Update_Test_Stage",                       @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Test_Stage)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_Stage_at_row, KeywordReportHeader.Test_Stage_at_col, header.Test_Stage);
            //}

            ////@"Update_Test_QTY_SN",                      @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Test_QTY_SN)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_QTY_SN_at_row, KeywordReportHeader.Test_QTY_SN_at_col, header.Test_QTY_SN);
            //}

            ////@"Update_Test_Period_Begin",                @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Test_Period_Begin)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_Period_Begin_at_row, KeywordReportHeader.Test_Period_Begin_at_col, header.Test_Period_Begin);
            //}

            ////@"Update_Test_Period_End",                  @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Test_Period_End)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_Period_End_at_row, KeywordReportHeader.Test_Period_End_at_col, header.Test_Period_End);
            //}

            ////@"Update_Judgement",                        @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Judgement)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col, header.Judgement);
            //}

            ////@"Update_Tested_by",                        @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Tested_by)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Tested_by_at_row, KeywordReportHeader.Tested_by_at_col, header.Tested_by);
            //}

            ////@"Update_Approved_by",                      @"true",
            //if (KeywordReport.DefaultKeywordReportHeader.Update_Approved_by)
            //{
            //    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Approved_by_at_row, KeywordReportHeader.Approved_by_at_col, header.Approved_by);
            //}
            return true;
        }

        static public Boolean UpdateKeywordReportHeader_full(Worksheet report_worksheet, TestReportOption header)
        {
            Boolean b_ret = false;
            try
            {
                //@"Update_Report_Title_by_Sheetname",        @"true",
                if (TestReport.Option.Update_Report_Title_by_Sheetname)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Title_at_row, TestReport.Title_at_col, header.Report_Title);
                    ExcelAction.SetFontColorToWhite(report_worksheet, 1, 1, 1, 1);  // A1 only
                }

                //@"Update_Model_Name",                       @"true",
                if (TestReport.Option.Update_Model_Name)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Model_Name_at_row, TestReport.Model_Name_at_col, header.Model_Name);
                }

                //@"Update_Part_No",                          @"true",
                if (TestReport.Option.Update_Part_No)
                {
                    String output_part_no = header.Part_No + "-" + header.Report_SheetName;
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Part_No_at_row, TestReport.Part_No_at_col, output_part_no);
                }

                //@"Update_Panel_Module",                     @"true",
                if (TestReport.Option.Update_Panel_Module)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Panel_Module_at_row, TestReport.Panel_Module_at_col, header.Panel_Module);
                }

                //@"Update_TCON_Board",                       @"true",
                if (TestReport.Option.Update_TCON_Board)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.TCON_Board_at_row, TestReport.TCON_Board_at_col, header.TCON_Board);
                }

                //@"Update_AD_Board",                         @"true",
                if (TestReport.Option.Update_AD_Board)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.AD_Board_at_row, TestReport.AD_Board_at_col, header.AD_Board);
                }

                //@"Update_Power_Board",                      @"true",
                if (TestReport.Option.Update_Power_Board)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Power_Board_at_row, TestReport.Power_Board_at_col, header.Power_Board);
                }

                //@"Update_Smart_BD_OS_Version",              @"true",
                if (TestReport.Option.Update_Smart_BD_OS_Version)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Smart_BD_OS_Version_at_row, TestReport.Smart_BD_OS_Version_at_col, header.Smart_BD_OS_Version);
                }

                //@"Update_Touch_Sensor",                     @"true",
                if (TestReport.Option.Update_Touch_Sensor)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Touch_Sensor_at_row, TestReport.Touch_Sensor_at_col, header.Touch_Sensor);
                }

                //@"Update_Speaker_AQ_Version",               @"true",
                if (TestReport.Option.Update_Speaker_AQ_Version)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Speaker_AQ_Version_at_row, TestReport.Speaker_AQ_Version_at_col, header.Speaker_AQ_Version);
                }

                //@"Update_SW_PQ_Version",                    @"true",
                if (TestReport.Option.Update_SW_PQ_Version)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.SW_PQ_Version_at_row, TestReport.SW_PQ_Version_at_col, header.SW_PQ_Version);
                }

                //@"Update_Test_Stage",                       @"true",
                if (TestReport.Option.Update_Test_Stage)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Test_Stage_at_row, TestReport.Test_Stage_at_col, header.Test_Stage);
                }

                //@"Update_Test_QTY_SN",                      @"true",
                if (TestReport.Option.Update_Test_QTY_SN)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Test_QTY_SN_at_row, TestReport.Test_QTY_SN_at_col, header.Test_QTY_SN);
                }

                //@"Update_Test_Period_Begin",                @"true",
                if (TestReport.Option.Update_Test_Period_Begin)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Test_Period_Begin_at_row, TestReport.Test_Period_Begin_at_col, header.Test_Period_Begin);
                }

                //@"Update_Test_Period_End",                  @"true",
                if (TestReport.Option.Update_Test_Period_End)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Test_Period_End_at_row, TestReport.Test_Period_End_at_col, header.Test_Period_End);
                }

                //@"Update_Judgement",                        @"true",
                if (TestReport.Option.Update_Judgement)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col, header.Judgement);
                }

                //@"Update_Tested_by",                        @"true",
                if (TestReport.Option.Update_Tested_by)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Tested_by_at_row, TestReport.Tested_by_at_col, header.Tested_by);
                }

                //@"Update_Approved_by",                      @"true",
                if (TestReport.Option.Update_Approved_by)
                {
                    ExcelAction.SetCellValue(report_worksheet, TestReport.Approved_by_at_row, TestReport.Approved_by_at_col, header.Approved_by);
                }

                b_ret = true;
            }
            catch (Exception ex)
            {
            }

            return b_ret;
        }

        //static public Boolean UpdateReportHeader(Worksheet ws, String Title = null, String SW_Version = null, String Test_Start = null, String Test_End = null,
        //                String Judgement = null, String Template = null)
        //{
        //    Boolean b_ret = false;
        //    // to-be-finished.
        //    if (Template != null)
        //    {

        //    }
        //    else
        //    {
        //        if (Title != null)
        //        {
        //            ExcelAction.SetCellValue(ws, KeywordReportHeader.Title_at_row, KeywordReportHeader.Title_at_col, Title);
        //        }
        //        if (SW_Version != null)
        //        {
        //            ExcelAction.SetCellValue(ws, KeywordReportHeader.SW_PQ_Version_at_row, KeywordReportHeader.SW_PQ_Version_at_col, Judgement);
        //        }
        //        if (Test_Start != null)
        //        {
        //            ExcelAction.SetCellValue(ws, KeywordReportHeader.Test_Period_Begin_at_row, KeywordReportHeader.Test_Period_Begin_at_col, Test_Start);
        //        }
        //        if (Test_End != null)
        //        {
        //            ExcelAction.SetCellValue(ws, KeywordReportHeader.Test_Period_End_at_row, KeywordReportHeader.Test_Period_End_at_col, Test_End);
        //        }
        //        if (Judgement != null)
        //        {
        //            ExcelAction.SetCellValue(ws, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col, Judgement);
        //        }
        //    }
        //    b_ret = true;
        //    return b_ret;
        //}

        //public static Boolean UpdateAllHeader(List<String> report_list, String Title = null, String SW_Version = null, String Test_Start = null, String Test_End = null,
        //                                String Judgement = null, String Template = null)
        //{
        //    // Create a temporary test plan to includes all files listed in List<String> report_filename
        //    List<TestPlan> do_plan = TestPlan.CreateTempPlanFromFileList(report_list);

        //    foreach (TestPlan plan in do_plan)
        //    {
        //        String path = Storage.GetDirectoryName(plan.ExcelFile);
        //        String filename = Storage.GetFileName(plan.ExcelFile);
        //        String sheet_name = plan.ExcelSheet;
        //        TestPlan.ExcelStatus test_plan_status;

        //        test_plan_status = plan.OpenDetailExcel(ReadOnly: false);
        //        if (test_plan_status == TestPlan.ExcelStatus.OK)
        //        {
        //            UpdateReportHeader(plan.TestPlanWorksheet, Title: Title, SW_Version: SW_Version, Test_Start: Test_Start,
        //                                    Test_End: Test_End, Judgement: Judgement, Template: Template);
        //            plan.SaveDetailExcel(plan.ExcelFile);
        //            plan.CloseDetailExcel();
        //        }
        //    }
        //    return true;
        //}

        static public Boolean HideKeywordResultBugRow(String excel_filename, Worksheet ws)
        {
            // Clear Keyword bug result and hide 2 rows (only Item row left)
            Boolean b_ret = false;
            try
            {
                TestPlan tp = TestPlan.CreateTempPlanFromFile(excel_filename);
                tp.TestPlanWorksheet = ws;
                tp.ExcelSheet = ws.Name;
                List<TestReportKeyword> keyword_list = ListKeyword_SingleReport(tp);
                foreach (TestReportKeyword keyword in keyword_list)
                {
                    double new_row_height = 0.2;
                    ExcelAction.Set_Row_Height(ws, keyword.BugListAtRow, new_row_height);
                    ExcelAction.Set_Row_Height(ws, keyword.BugStatusAtRow, new_row_height);
                }

                b_ret = true;
            }
            catch (Exception ex)
            {
            }
            return b_ret;
        }

        static public Boolean ClearKeywordBugResult(String excel_filename, Worksheet ws)
        {
            Boolean b_ret = false;
            try
            {
                TestPlan tp = TestPlan.CreateTempPlanFromFile(excel_filename);
                tp.TestPlanWorksheet = ws;
                tp.ExcelSheet = ws.Name;
                List<TestReportKeyword> keyword_list = ListKeyword_SingleReport(tp);
                foreach (TestReportKeyword keyword in keyword_list)
                {
                    ExcelAction.SetCellValue(ws, keyword.ResultAtRow, keyword.ResultAtColumn, " ");
                    int temp_col = keyword.BugStatusAtColumn;
                    ExcelAction.SetCellValue(ws, keyword.BugStatusAtRow, temp_col++, " ");
                    ExcelAction.SetCellValue(ws, keyword.BugStatusAtRow, temp_col++, " ");
                    ExcelAction.SetCellValue(ws, keyword.BugStatusAtRow, temp_col++, " ");
                    ExcelAction.SetCellValue(ws, keyword.BugStatusAtRow, temp_col++, " ");
                    ExcelAction.SetCellValue(ws, keyword.BugStatusAtRow, temp_col++, " ");
                    ExcelAction.SetCellValue(ws, keyword.BugListAtRow, keyword.BugListAtColumn, " ");
                    double new_row_height = (StyleString.default_size + 1) * 2 * 0.75;
                    //ExcelAction.Unhide_Row(ws, keyword.BugListAtRow);
                    ExcelAction.Set_Row_Height(ws, keyword.BugListAtRow, new_row_height);
                    ExcelAction.Set_Row_Height(ws, keyword.BugStatusAtRow, new_row_height);
                }

                b_ret = true;
            }
            catch (Exception ex)
            {
            }
            return b_ret;
        }

        static public Boolean ClearReportBugCount(Worksheet ws)
        {
            Boolean b_ret = false;
            try
            {
                ExcelAction.SetCellValue(ws, PassCnt_at_row, PassCnt_at_col, " ");
                ExcelAction.SetCellValue(ws, FailCnt_at_row, FailCnt_at_col, " ");
                ExcelAction.SetCellValue(ws, ConditionalPass_string_at_row, ConditionalPass_string_at_col, CONDITIONAL_PASS_str + ":");
                ExcelAction.SetCellValue(ws, ConditionalPassCnt_at_row, ConditionalPassCnt_at_col, " ");
                b_ret = true;
            }
            catch (Exception ex)
            {
            }
            return b_ret;
        }

        static public Boolean ClearJudgement(Worksheet ws)
        {
            Boolean b_ret = false;
            try
            {
                ExcelAction.SetCellValue(ws, TestReport.Judgement_at_row, TestReport.Judgement_at_col, " ");
                b_ret = true;
            }
            catch (Exception ex)
            {
            }
            return b_ret;
        }

        static public Boolean UpdateSampleSN_to_common_string(Worksheet ws)
        {
            Boolean b_ret = UpdateSampleSN(ws, Option.SampleSN_String);
            return b_ret;
        }

        static public Boolean UpdateSampleSN(Worksheet ws, String new_sample_sn)
        {
            Boolean b_ret = false;
            int SN_row = 0, SN_title_col = ExcelAction.ColumnNameToNumber("B"), SN_number_col = ExcelAction.ColumnNameToNumber("C");

            // Find SN title
            int conclusion_start_row = row_test_brief_start, conclusion_end_row = row_test_brief_end;
            for (int row_index = (row_test_brief_end); row_index <= (row_test_brief_end + 3); row_index++)
            {
                String text = ExcelAction.GetCellTrimmedString(ws, row_index, SN_title_col);
                if (CheckIfStringMeetsSampleSN(text))
                {
                    SN_row = row_index;
                    break;
                }
            }

            // If not found, insert it 2 rows below conclusion
            if (SN_row == 0)
            {
                SN_row = FindTitle_Conclusion(ws) + 2;
                // insert a new one
                ExcelAction.Insert_Row(ws, SN_row);

                String SN_Title = "Sample S/N:";
                StyleString StyleString_SN_Title = new StyleString(SN_Title);
                StyleString_SN_Title.FontStyle = FontStyle.Bold;
                //ExcelAction.SetCellValue(ws, SN_row, SN_title_col, SN_Title);
                StyleString.WriteStyleString(ws, SN_row, SN_title_col, StyleString_SN_Title.ConvertToList());
            }

            // Update SN string
            try
            {
                int col_start = SN_number_col,
                    col_end = ExcelAction.ColumnNameToNumber("M");

                ExcelAction.ClearContent(ws, SN_row, SN_number_col + 1, SN_row, col_end);
                ExcelAction.CellTextAlignLeft(ws, SN_row, SN_number_col);
                
                //ExcelAction.SetCellValue(ws, SN_row, SN_number_col, new_sample_sn);
                //String SN_Font = XMLConfig.ReadAppSetting_String("SampleSN_String_Font");
                //int SN__FontSize = XMLConfig.ReadAppSetting_int("SampleSN_String_FontSize");
                //Color SN_FontColor = XMLConfig.ReadAppSetting_Color("SampleSN_String_FontColor");
                //FontStyle SN_FontStyle = XMLConfig.ReadAppSetting_FontStyle("SampleSN_String_FontStyle");
                StyleString style_string_new_sample_sn = new StyleString(new_sample_sn, Option.SN_FontColor, Option.SN_Font, Option.SN_FontSize, Option.SN_FontStyle);
                StyleString.WriteStyleString(ws, SN_row, SN_number_col, style_string_new_sample_sn.ConvertToList());

                ExcelAction.Merge(ws, SN_row, col_start, SN_row, col_end);

                b_ret = true;
            }
            catch (Exception ex)
            {
            }
            return b_ret;
        }

        // Assumption:
        //      1. Bug / TC has been opened and processed
        //      2. Current worksheet is the report worksheet
        //      3. If there isn't corresponding testcase, judgement/conclusion is set to default
        static public Boolean Update_Conclusion_Judgement_by_linked_issue(Worksheet worksheet)
        {
            Boolean b_ret = false;

            String judgement_str;
            List<StyleString> linked_issue_description_on_this_report;

            b_ret = GetSortedFilteredLinkIssueAndJudgementString(worksheet, out linked_issue_description_on_this_report, out judgement_str);

            // update conclusion
            b_ret = ReplaceConclusionWithBugList(worksheet, linked_issue_description_on_this_report);

            // update judgement
            ExcelAction.CellActivate(worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col);
            ExcelAction.SetCellValue(worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col, judgement_str);

            return b_ret;
        }

        static public Boolean Update_Conclusion_only_by_linked_issue(Worksheet worksheet)
        {
           
            Boolean b_ret = false;
            String judgement_str;           // obtained but not use in this function
            List<StyleString> linked_issue_description_on_this_report;

            b_ret = GetSortedFilteredLinkIssueAndJudgementString(worksheet, out linked_issue_description_on_this_report, out judgement_str);

            // update conclusion
            b_ret = ReplaceConclusionWithBugList(worksheet, linked_issue_description_on_this_report);

            return b_ret;
        }

        static public Boolean Update_Judgement_only_by_linked_issue(Worksheet worksheet)
        {
            Boolean b_ret = false;

            String judgement_str;
            List<StyleString> linked_issue_description_on_this_report;

            b_ret = GetSortedFilteredLinkIssueAndJudgementString(worksheet, out linked_issue_description_on_this_report, out judgement_str);

            // update judgement
            ExcelAction.CellActivate(worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col);
            ExcelAction.SetCellValue(worksheet, TestReport.Judgement_at_row, TestReport.Judgement_at_col, judgement_str);

            return b_ret;
        }

        //
        // return: false if corresponding report is not available (so returning default for both Return_Issue_List & Judgement_String)
        // out: Return_Issue_List - sorted and filtered linked issue list
        //      Judgement_String - judgement according to sorted/filtered linked issue
        //
        static public Boolean GetSortedFilteredLinkIssueAndJudgementString(Worksheet worksheet, out List<StyleString> Return_Issue_List, out String Judgement_String)
        {
            Boolean b_ret = false;

            String sheet_name = worksheet.Name;
            String judgement_str = TestReport.TestReport_Default_Judgement; 
            StyleString default_conclusion = new StyleString( TestReport.TestReport_Default_Conclusion, ReportGenerator.LinkIssue_report_FontColor, 
                            ReportGenerator.LinkIssue_report_Font, ReportGenerator.LinkIssue_report_FontSize);
            List<StyleString> linked_issue_description_on_this_report = default_conclusion.ConvertToList();      

            //judgement_str = ExcelAction.GetCellTrimmedString(worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col);

            // Find the TC meets the sheet-name
            if (ReportGenerator.GetTestcaseLUT_by_Sheetname().ContainsKey(sheet_name))          // if TC is available
            {
                // key string of all linked issues
                String links = ReportGenerator.GetTestcaseLUT_by_Sheetname()[sheet_name].Links;
                // key string to List of Issue
                List<Issue> linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                // List of Issue filtered by status
                List<Issue> filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.List_of_status_to_filter_for_tc_linked_issue);
                // Sort issue by Severity and Key valie
                List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
                // Convert list of sorted linked issue to description list
                linked_issue_description_on_this_report = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);

                // decide judgement result based on linked issue severity and count
                judgement_str = Judgement_Decision_by_Linked_Issue_List(linked_issue_list);

                b_ret = true;
            }
            else
            {
                // return default if no corresponding TC
                LogMessage.WriteLine("Worksheet " + sheet_name + " not found in GetSortedFilteredLinkIssueAndJudgementString()");
                //judgement_str = " ";
                //linked_issue_description_on_this_report.Clear();
                //linked_issue_description_on_this_report.Add(KeywordReportHeader.blank_space);
                b_ret = false;
            }
            // get data only here so no replace
            //b_ret = ReplaceConclusionWithBugList(worksheet, linked_issue_description_on_this_report);

            Judgement_String = judgement_str;
            Return_Issue_List = linked_issue_description_on_this_report;
            return b_ret;

        }

        // This function is used to get judgement result (only read and no update to report) of keyword report
        //static public Boolean GetJudgementValue(String report_workbook, String report_worksheet, out String judgement_str)
        static public Boolean GetJudgementPurposeCriteriaValue(String report_workbook, String report_worksheet, out String judgement_str, out String purpose_str, out String criteria_str)
        {
            Boolean b_ret = false;
            String ret_str = ""; // default returning judgetment_str;
            purpose_str = criteria_str = "";

            // 1. Open Excel and find the sheet
            // File exist check is done outside
            Workbook wb_judgement = ExcelAction.OpenExcelWorkbook(report_workbook);
            if (wb_judgement == null)
            {
                LogMessage.WriteLine("ERR: Open workbook in GetJudgementValue: " + report_workbook);
                judgement_str = ret_str;
                b_ret = false;
            }
            else
            {
                // 2 Open worksheet
                Worksheet ws_judgement = ExcelAction.Find_Worksheet(wb_judgement, report_worksheet);
                if (ws_judgement == null)
                {
                    LogMessage.WriteLine("ERR: Open worksheet in GetJudgementValue: " + report_workbook + " sheet: " + report_worksheet);
                    judgement_str = ret_str;
                    b_ret = false;
                }
                else
                {
                    // 3. Get Judgement value
                    Object obj = ExcelAction.GetCellValue(ws_judgement, TestReport.Judgement_at_row, TestReport.Judgement_at_col);
                    if (obj != null)
                    {
                        judgement_str = (String)obj;
                        b_ret = true;
                    }
                    else
                    {
                        judgement_str = ret_str;
                        b_ret = false;
                    }

                    // 3.3.0: store text of purpose & criteia for updating into TC summary report
                    int search_start_row = 10, search_end_row = 19;
                    for (int row_index = search_start_row; row_index <= search_end_row; row_index++)
                    {
                        String text = ExcelAction.GetCellTrimmedString(ws_judgement, row_index, col_indentifier);
                        if (CheckIfStringMeetsPurpose(text))
                        {
                            row_index++;
                            purpose_str = ExcelAction.GetCellTrimmedString(ws_judgement, row_index, ExcelAction.ColumnNameToNumber('C'));
                            continue;
                        }
                        else if (CheckIfStringMeetsCriteria(text))
                        {
                            row_index++;
                            criteria_str = ExcelAction.GetCellTrimmedString(ws_judgement, row_index, ExcelAction.ColumnNameToNumber('C'));
                            continue;
                        }
                    }
                }

                // Close excel if open succeeds
                ExcelAction.CloseExcelWorkbook(wb_judgement);
            }
            return b_ret;
        }

        // This function is used to get judgement result (only read and no update to report) of test report
        //static public Boolean GetAllKeywordIssueOnReport(String report_filename, String report_sheetname, out StyleString issue_list_str)
        //{
        //    Boolean b_ret = false;
        //    StyleString ret_str = new StyleString(); 

        //    // 1. Open Excel and find the sheet
        //    // File exist check is done outside
        //    Workbook wb_report = ExcelAction.OpenExcelWorkbook(report_filename);
        //    if (wb_report == null)
        //    {
        //        ConsoleWarning("ERR: Open workbook in " + System.Reflection.MethodBase.GetCurrentMethod().Name + ": " + report_filename);
        //        issue_list_str = ret_str;
        //        b_ret = false;
        //    }
        //    else
        //    {
        //        // 2 Open worksheet
        //        Worksheet ws_report = ExcelAction.Find_Worksheet(wb_report, report_sheetname);
        //        if (ws_report == null)
        //        {
        //            ConsoleWarning("ERR: Open worksheet in " + System.Reflection.MethodBase.GetCurrentMethod().Name + ": " + report_filename + " sheet: " + report_sheetname);
        //            issue_list_str = ret_str;
        //            b_ret = false;
        //        }
        //        else
        //        {
        //            TestPlan report_testplan = TestPlan.CreateTempPlanFromFile(report_filename);

        //            // 3. Get Keyword issue list
        //            List<TestPlanKeyword> keyword_report = KeywordReport.ListKeyword_SingleReport(report_testplan);

        //            foreach (TestPlanKeyword tp_keyword in keyword_report)
        //            {
        //                int row = tp_keyword.BugListAtRow, col = tp_keyword.BugListAtColumn;
        //            }

        //            //List<TestPlanKeyword> ws_keyword_list = KeywordReport.FilterSingleReportKeyword(keyword_list, report_workbook, report_worksheet);


        //            Object obj = ExcelAction.GetCellValue(ws_report, TestReport.Judgement_at_row, TestReport.Judgement_at_col);
        //            if (obj != null)
        //            {
        //                //issue_list_str = obj;
        //                b_ret = true;
        //            }
        //            else
        //            {
        //                issue_list_str = ret_str;
        //                b_ret = false;
        //            }
        //        }

        //        // Close excel if open succeeds
        //        ExcelAction.CloseExcelWorkbook(wb_report);
        //    }
        //    return b_ret;
        //}

        static public String DecideDestinationFilename(String src_dir, String dest_dir, String full_filename)
        {
            String ret_str;

            if ((dest_dir == "") || !Storage.DirectoryExists(src_dir))
            {
                ret_str = Storage.GenerateFilenameWithDateTime(full_filename);
            }
            else
            {
                ret_str = full_filename.Replace(src_dir, dest_dir);
            }

            return ret_str;
        }

        // 
        // Input: Standard Test Report main file
        // Output: keyword list of all "Do" test-plans
        //
        static public List<TestReportKeyword> ListAllDetailedTestPlanKeywordTask(String report_root_dir, String output_filename)
        {
            // Clear keyword log report data-table
            ReportGenerator.excel_not_report_log.Clear();
            // 0.1 List all files under report_root_dir.
            List<String> file_list = Storage.ListFilesUnderDirectory(report_root_dir);
            // 0.2 filename check to exclude non-report files.
            List<String> report_filename = Storage.FilterFilename(file_list);
            // 0.3 output files in file_list but not in report_filename into Not_Keyword_File
            foreach (String report_file in report_filename)
            {
                file_list.Remove(report_file);
            }
            foreach (String NG_file in file_list)
            {
                String path, filename;
                path = Storage.GetDirectoryName(NG_file);
                filename = Storage.GetFileName(NG_file);
                ReportFileRecord nrfr_item = new ReportFileRecord(path, filename);
                nrfr_item.SetFlagFail(excelfilenamefail: true);
                ReportGenerator.excel_not_report_log.Add(nrfr_item);
            }

            //
            // 1. Create a temporary test plan (do_plan) to include all report files 
            //
            // 1.1 Init an empty plan
            List<TestPlan> do_plan = new List<TestPlan>();

            // 1.2 Create a temporary test plan to includes all files listed in List<String> report_filename
            do_plan = TestPlan.CreateTempPlanFromFileList(report_filename);

            //
            // 2. Search keywords within all selected file (2.1) and use those keywords to find out issues containing keywords.
            //
            // 2.1. Find keyword for all selected file (as listed in temprary test plan)
            //
            List<TestReportKeyword> keyword_list = ListAllKeyword(do_plan);

            // Output keyword list log excel here.
            KeyWordListReport.OutputKeywordLog(report_root_dir, keyword_list, ReportGenerator.excel_not_report_log, output_filename);


            // Output updated report with recommended sheetname.
            if (TestReport.KeywordIssue_Auto_Correct_Sheetname == true)
            {
                String dest_dir = Storage.GenerateDirectoryNameWithDateTime(report_root_dir);
                // ReportGenerator.excel_not_report_log
                foreach (ReportFileRecord item in ReportGenerator.excel_not_report_log)
                {
                    String path, filename, expected_sheetname;
                    Boolean excelfilenameOK, openfileOK, findWorksheetOK, findAnyKeyword, otherFailure;

                    item.GetRecord(out path, out filename, out expected_sheetname, out excelfilenameOK, out openfileOK, out findWorksheetOK,
                            out findAnyKeyword, out otherFailure);

                    if ((openfileOK == true) && (findWorksheetOK == false) && (otherFailure == false))
                    {
                        String full_filename = Storage.GetValidFullFilename(path, filename);
                        // Open Excel and find the sheet
                        Workbook wb = ExcelAction.OpenExcelWorkbook(filename: full_filename, ReadOnly: false);
                        if (wb == null)
                        {
                            LogMessage.WriteLine("ERR: Open workbook in auto-correct-worksheet-name of ListAllDetailedTestPlanKeywordTask(): " + full_filename);
                            continue;
                        }

                        // Use first worksheet and rename it.
                        Worksheet ws = wb.Sheets[1];
                        ws.Name = expected_sheetname;

                        // Save the updated report file file to either 
                        //  (1) filename with yyyyMMddHHmmss if dest_dir is not specified
                        //  (2) the same filename but to the sub-folder of same strucure under new root-folder "dest_dir"
                        String dest_filename = DecideDestinationFilename(report_root_dir, dest_dir, full_filename);
                        String dest_filename_dir = Storage.GetDirectoryName(dest_filename);
                        // if parent directory does not exist, create recursively all parents
                        Storage.CreateDirectory(dest_filename_dir, auto_parent_dir: true);
                        ExcelAction.CloseExcelWorkbook(wb, SaveChanges: true, AsFilename: dest_filename);
                    }
                }
            }

            return keyword_list;
        }

        static public Dictionary<String, List<TestReportKeyword>> GenerateKeywordLUT_by_Sheetname(List<TestReportKeyword> keyword_list)
        {
            Dictionary<String, List<TestReportKeyword>> ret_dic = new Dictionary<String, List<TestReportKeyword>>();

            foreach (TestReportKeyword tpk in keyword_list)
            {
                String sheet = tpk.Worksheet;

                // if key (current) exists in dictionary, add new dictionary pair (keyword, sheetname-list) with value is empty List.
                if (ret_dic.ContainsKey(sheet) == false)
                {
                    ret_dic.Add(sheet, new List<TestReportKeyword>());
                }
                // Add item (sheet) into the list of dictionary pair (keyword, sheetname-list)
                ret_dic[sheet].Add(tpk);
            }
            return ret_dic;
        }

        static public Boolean AutoCorrectReport_SingleFile(String source_file, String destination_file, Workbook wb_template, Boolean always_save = false)
        {
            TestReportOption out_header; // not used for this version of API.
            return AutoCorrectReport_SingleFile(source_file, destination_file, wb_template, out out_header, always_save);
        }

        // Copy and update worksheet name & header & bug-result (configurable) -- to be used for starting a new project based on reports from anywhere
        static public Boolean AutoCorrectReport_SingleFile(String source_file, String destination_file, Workbook wb_template, out TestReportOption out_header, Boolean always_save = false)
        {
            Boolean file_has_been_updated = false;
            out_header = new TestReportOption();

            destination_file = Storage.GetFullPath(destination_file);
            if (Storage.IsReportFilename(destination_file) == false)
            {
                // Do nothing if new filename does not look like a report filename.
                return file_has_been_updated;
            }

            // Open Excel workbook
            source_file = Storage.GetFullPath(source_file);
            Workbook wb = ExcelAction.OpenExcelWorkbook(filename: source_file, ReadOnly: false);
            if (wb == null)
            {
                LogMessage.WriteLine("ERR: Open workbook in AutoCorrectReport_SingleFile(): " + source_file);
                return false;
            }

            // If valid sheet_name does not exist, use first worksheet .
            String current_sheet_name = TestPlan.GetSheetNameAccordingToFilename(source_file);
            TestReport.Option.Report_SheetName = current_sheet_name;
            Worksheet ws;
            if (ExcelAction.WorksheetExist(wb, current_sheet_name) == false)
            {
                ws = wb.Sheets[1];
            }
            else
            {
                ws = ExcelAction.Find_Worksheet(wb, current_sheet_name);
            }

            // Update sheetname (when the option is true)
            if (TestReport.Option.Report_C_Update_Report_Sheetname)
            {
                String new_sheet_name = TestPlan.GetSheetNameAccordingToFilename(destination_file);
                ws.Name = new_sheet_name;
                TestReport.Option.Report_SheetName = new_sheet_name;
                file_has_been_updated = true;
            }

            //Report_C_Update_Header_by_Template
            if (TestReport.Option.Report_C_Update_Header_by_Template == true)
            {
                if (ExcelAction.WorksheetExist(wb_template, HeaderTemplate.SheetName_HeaderTemplate))
                {
                    Worksheet ws_template = ExcelAction.Find_Worksheet(wb_template, HeaderTemplate.SheetName_HeaderTemplate);
                    String filename = TestPlan.GetReportTitleAccordingToFilename(destination_file);
                    String sheetname = ws.Name;
                    HeaderTemplate.UpdateVariables_FilenameSheetname(filename: filename, sheetname: sheetname);
                    //HeaderTemplate.CopyAndUpdateHeader(ws_template, ws);
                    HeaderTemplate.CopyAndUpdateHeader_with_KEEP(ws_template, ws);
                }

                //Report_C_Replace_Conclusion
                if (TestReport.Option.Report_C_Replace_Conclusion == true)
                {
                    //StyleString blank_space = new StyleString(" ", StyleString.default_color, StyleString.default_font, StyleString.default_size);
                    TestReport.ReplaceConclusionWithBugList(ws, TestReport.blank_space_list);
                    file_has_been_updated = true;
                }

            }
            else
            {
                // Update header (when the option is true)
                if (TestReport.Option.Report_C_Update_Full_Header == true)
                {
                    String new_title = TestPlan.GetReportTitleAccordingToFilename(destination_file);
                    TestReport.Option.Report_Title = new_title;
                    // sheet-name is not defined as part of header --> it should be part of excel report (eg. filename, sheetname)
                    //KeywordReport.DefaultKeywordReportHeader.Report_SheetName = new_sheet_name;
                    TestReport.UpdateKeywordReportHeader_full(ws, TestReport.Option);
                    file_has_been_updated = true;
                }

                //Report_C_Replace_Conclusion
                if (TestReport.Option.Report_C_Replace_Conclusion == true)
                {
                    //StyleString blank_space = new StyleString(" ", StyleString.default_color, StyleString.default_font, StyleString.default_size);
                    TestReport.ReplaceConclusionWithBugList(ws, TestReport.blank_space_list);
                    file_has_been_updated = true;
                }

                // Clear bug-list, bug-count, Pass/Fail/Conditional_Pass count, judgement
                if (TestReport.Option.Report_C_Clear_Keyword_Result)
                {
                    TestReport.ClearKeywordBugResult(source_file, ws);
                    TestReport.ClearReportBugCount(ws);
                    TestReport.ClearJudgement(ws);
                    file_has_been_updated = true;
                }
            }

            // Hide keyword result/bug-list row -- after clear because it is un-hide after clear
            if (TestReport.Option.Report_C_Hide_Keyword_Result_Bug_Row)
            {
                TestReport.HideKeywordResultBugRow(source_file, ws);
                file_has_been_updated = true;
            }

            if (TestReport.Option.Report_C_Update_Conclusion)
            {
                TestReport.Update_Conclusion_only_by_linked_issue(ws);
                file_has_been_updated = true;
            }

            if (TestReport.Option.Report_C_Update_Judgement)
            {
                TestReport.Update_Judgement_only_by_linked_issue(ws);
                file_has_been_updated = true;
            }

            if (TestReport.Option.Report_C_Update_Sample_SN)
            {
                TestReport.UpdateSampleSN_to_common_string(ws);
                file_has_been_updated = true;
            }

            if (TestReport.Option.Report_C_Remove_AUO_Internal)
            {
                // step 1: remove sheets which are not to be released
                String sheet_name_to_keep = ws.Name;
                if (wb.Sheets.Count > 1)
                {
                    // work-sheet can be deleted only when there are two or more sheets
                    for (int sheet_index = wb.Sheets.Count; sheet_index > 0; sheet_index--)
                    {
                        String temp_sheet_name = wb.Sheets[sheet_index].Name;
                        if (temp_sheet_name.Length >= sheet_name_to_keep.Length)
                        {
                            if (temp_sheet_name.Substring(0, sheet_name_to_keep.Length) == sheet_name_to_keep)
                            {
                                continue;
                            }
                        }
                        wb.Sheets[sheet_index].Delete();
                        file_has_been_updated = true;
                    }
                }
                // step 2: remove contents on the sheet which are not to be released
                if (TestReport.Option.Report_C_Remove_AUO_Internal_remove_Method)
                {

                    //CheckIfStringMeetsMethod
                    int search_start_row = TestReport.row_test_brief_start, search_end_row = TestReport.row_test_brief_end;
                    for (int row_index = search_start_row; row_index <= search_end_row; row_index++)
                    {
                        String text = ExcelAction.GetCellTrimmedString(wb.Sheets[1], row_index, TestReport.col_indentifier);
                        if (TestReport.CheckIfStringMeetsMethod(text))
                        {
                            ExcelAction.ClearContent(wb.Sheets[1], row_index, 1, row_index + 1, TestReport.col_default_report_right_border);
                            double new_row_height = 0.2;
                            ExcelAction.Set_Row_Height(wb.Sheets[1], row_index, new_row_height);
                            ExcelAction.Set_Row_Height(wb.Sheets[1], row_index + 1, new_row_height);
                            file_has_been_updated = true;
                            break;
                        }
                    }
                }
            }

            if ((file_has_been_updated) || (always_save))
            {
                // Something has been updated or always save (ex: to copy file & update) ==> save to excel file
                String destination_dir = Storage.GetDirectoryName(destination_file);
                // if parent directory does not exist, create recursively all parents
                if (Storage.DirectoryExists(destination_dir) == false)
                {
                    Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
                }
                ExcelAction.SaveExcelWorkbook(wb, filename: destination_file);
            }
            else
            {
                // Doing nothing here.
            }
            // Close Excel workbook
            ExcelAction.CloseExcelWorkbook(wb);

            return file_has_been_updated;
        }

        /*
        static public Boolean AutoCorrectReport_by_Folder(String report_root, String Output_dir)
        {
            Boolean b_ret = false;

            // 0.1 List all files under report_root_dir.
            List<String> file_list = Storage.ListFilesUnderDirectory(report_root);
            // 0.2 filename check to exclude non-report files.
            List<String> report_filename = Storage.FilterFilename(file_list);

            foreach (String source_report in report_filename)
            {
                String dest_filename = TestReport.DecideDestinationFilename(report_root, Output_dir, source_report); // replace folder name
                b_ret |= AutoCorrectReport_SingleFile(source_file: source_report, destination_file: dest_filename, wb_template: new Workbook(), always_save: true);
            }

            return b_ret;
        }
        */

        //static public Boolean AutoCorrectReport(String report_root, String Output_dir = "")
        //{
        //    Boolean b_ret = false;

        //    // 0.1 List all files under report_root_dir.
        //    List<String> file_list = Storage.ListFilesUnderDirectory(report_root);
        //    // 0.2 filename check to exclude non-report files.
        //    List<String> report_filename = Storage.FilterFilename(file_list);

        //    //Output_dir = Storage.GetFullPath(Output_dir);

        //    // 1.1 Init an empty plan
        //    List<TestPlan> do_plan = new List<TestPlan>();

        //    // 1.2 Create a temporary test plan to includes report_file
        //    do_plan = TestPlan.CreateTempPlanFromFileList(report_filename);
        //    Boolean output_to_different_path = ((Output_dir=="")||(report_root==Output_dir))?false:true;

        //    foreach (TestPlan plan in do_plan)
        //    {
        //        String path = Storage.GetDirectoryName(plan.ExcelFile);
        //        String filename = Storage.GetFileName(plan.ExcelFile);
        //        String full_filename = Storage.GetFullPath(plan.ExcelFile);
        //        String sheet_name = plan.ExcelSheet;
        //        Boolean file_has_been_updated = output_to_different_path;

        //        // Open Excel workbood
        //        Workbook wb = ExcelAction.OpenExcelWorkbook(filename: full_filename, ReadOnly: false);
        //        if (wb == null)
        //        {
        //            ConsoleWarning("ERR: Open workbook in AutoCorrectReport(): " + full_filename);
        //            continue;
        //        }

        //        // If valid sheet_name does not exist, use first worksheet and rename it.
        //        Worksheet ws;
        //        if (ExcelAction.WorksheetExist(wb, sheet_name) == false)
        //        {
        //            ws = wb.Sheets[1];
        //            ws.Name = sheet_name;
        //            file_has_been_updated = true;
        //        }
        //        else
        //        {
        //            ws = ExcelAction.Find_Worksheet(wb, sheet_name);
        //        }

        //        // Update header 
        //        String new_title = TestPlan.GetReportTitleAccordingToFilename(filename);
        //        String existing_title = ExcelAction.GetCellTrimmedString(ws, TestReport.Title_at_row, TestReport.Title_at_col);
        //        if (existing_title != new_title)
        //        {
        //            TestReport.UpdateReportHeader(ws,Title: new_title);
        //            file_has_been_updated = true;
        //        }

        //        if (file_has_been_updated)
        //        {
        //            String dest_filename;
        //            // Something has been updated, save to excel file

        //            dest_filename = DecideDestinationFilename(report_root, Output_dir, full_filename);
        //            String dest_filename_dir = Storage.GetDirectoryName(dest_filename);
        //            // if parent directory does not exist, create recursively all parents
        //            Storage.CreateDirectory(dest_filename_dir, auto_parent_dir: true);
        //            ExcelAction.SaveExcelWorkbook(wb, filename: dest_filename);
        //            b_ret = true;
        //        }
        //        ExcelAction.CloseExcelWorkbook(wb);
        //    }
        //    return b_ret;
        //}

        //static public ReadReport(String root_dir)
        //{
        //}
    }
}
