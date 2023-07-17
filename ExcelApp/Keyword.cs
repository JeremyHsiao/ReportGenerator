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

    public class TestPlanKeyword
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

        public TestPlanKeyword() { TestPlanKeywordInit(); }
        public TestPlanKeyword(String Keyword, String Workbook = "", String Worksheet = "", int AtRow = 0, int AtColumn = 0,
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
        public int ResultListAtRow   // propertyd
        {
            get { return result_at_row; }   // get method
            set { result_at_row = value; }  // set method
        }
        public int ResultListAtColumn   // property
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

        public List<String> IssueList   // property
        {
            get { return issue_list; }   // get method
            set { issue_list = value; }  // set method
        }

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

        public void UpdateIssueList()
        {
            List<String> ret_str = new List<String>();

            if (keyword != "")
            {
            }

            IssueList = ret_str;
        }

        public void UpdateIssueDescriptionList(List<StyleString> description)
        {
            List<StyleString> ret_str = new List<StyleString>();

            if (IssueList == null)
            {
                UpdateIssueList();
            }

            if (IssueList != null)
            {

            }

            IssueDescriptionList = ret_str;
        }

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

        public IssueCount Calculate_Issue()
        {
            IssueCount ret_ic = new IssueCount();
            foreach (Issue issue in this.KeywordIssues)
            {
                if (issue.Status == Issue.STR_CLOSE)
                {
                    switch (issue.Severity[0])
                    {
                        case 'A':
                            ret_ic.Closed_A++;
                            break;
                        case 'B':
                            ret_ic.Closed_B++;
                            break;
                        case 'C':
                            ret_ic.Closed_C++;
                            break;
                        case 'D':
                            ret_ic.Closed_D++;
                            break;
                    }
                }
                else if (issue.Status == Issue.STR_WAIVE)
                {
                    switch (issue.Severity[0])
                    {
                        case 'A':
                            ret_ic.Waived_A++;
                            break;
                        case 'B':
                            ret_ic.Waived_B++;
                            break;
                        case 'C':
                            ret_ic.Waived_C++;
                            break;
                        case 'D':
                            ret_ic.Waived_D++;
                            break;
                    }
                }
                else // if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
                {
                    switch (issue.Severity[0])
                    {
                        case 'A':
                            ret_ic.Severity_A++;
                            break;
                        case 'B':
                            ret_ic.Severity_B++;
                            break;
                        case 'C':
                            ret_ic.Severity_C++;
                            break;
                        case 'D':
                            ret_ic.Severity_D++;
                            break;
                    }

                }
            }
            return ret_ic;
        }
    }

    public class KeywordReportHeader
    {
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

        public static int Title_at_row = 1, Title_at_col = ('A' - 'A' + 1);

        public static int Period_Start_at_row = 8, Period_Start_at_col = ('L' - 'A' + 1);
        public static int Period_End_at_row = 8, Period_End_at_col = ('M' - 'A' + 1);
        //        public static int Judgement_at_row = 9, Judgement_at_col = ('D' - 'A' + 1);
        public static int Judgement_string_at_row = 9, Judgement_string_at_col = 2;

        public static int Model_Name_at_row = 3, Model_Name_at_col = ('D' - 'A' + 1);
        public static int Part_No_at_row = 3, Part_No_at_col = ('J' - 'A' + 1);

        public static int Panel_Module_at_row = 4, Panel_Module_at_col = ('D' - 'A' + 1);
        public static int TCON_Board_at_row = 4, TCON_Board_at_col = ('J' - 'A' + 1);

        public static int AD_Board_at_row = 5, AD_Board_at_col = ('D' - 'A' + 1);
        public static int Power_Board_at_row = 5, Power_Board_at_col = ('J' - 'A' + 1);

        public static int Smart_BD_OS_Version_at_row = 6, Smart_BD_OS_Version_at_col = ('D' - 'A' + 1);
        public static int Touch_Sensor_at_row = 6, Touch_Sensor_at_col = ('J' - 'A' + 1);

        public static int Speaker_AQ_Version_at_row = 7, Speaker_AQ_Version_at_col = ('D' - 'A' + 1);
        public static int SW_PQ_Version_at_row = 7, SW_PQ_Version_at_col = ('J' - 'A' + 1);

        public static int Test_Stage_at_row = 8, Test_Stage_at_col = ('D' - 'A' + 1);
        public static int Test_QTY_SN_at_row = 8, Test_QTY_SN_at_col = ('H' - 'A' + 1);
        public static int Test_Period_Begin_at_row = 8, Test_Period_Begin_at_col = ('L' - 'A' + 1);
        public static int Test_Period_End_at_row = 8, Test_Period_End_at_col = ('M' - 'A' + 1);

        public static int Judgement_at_row = 9, Judgement_at_col = ('D' - 'A' + 1);
        public static int Tested_by_at_row = 9, Tested_by_at_col = ('H' - 'A' + 1);
        public static int Approved_by_at_row = 9, Approved_by_at_col = ('L' - 'A' + 1);

        //public static int Part_No_at_row = 3, Part_No_at_col = ('J' - 'A' + 1);
        //public static int SW_Version_at_row = 7, SW_Version_at_col = ('J' - 'A' + 1);
        //public static int Period_Start_at_row = 8, Period_Start_at_col = ('L' - 'A' + 1);
        //public static int Period_End_at_row = 8, Period_End_at_col = ('M' - 'A' + 1);
        //public static int Judgement_at_row = 9, Judgement_at_col = ('D' - 'A' + 1);
        //public static int Judgement_string_at_row = 9, Judgement_string_at_col = 2;

    }

    public static class KeywordReport
    {
        static public String TestReport_Default_Output_Dir = "";

        static private void ConsoleWarning(String function, int row, int col)
        {
            Console.WriteLine("Warning: please check " + function + " at (" + row.ToString() + "," + col.ToString() + ")");
        }
        static private void ConsoleWarning(String function, int row)
        {
            Console.WriteLine("Warning: please check " + function + " at line " + row.ToString());
        }
        static private void ConsoleWarning(String function)
        {
            Console.WriteLine("Warning: please check " + function);
        }

        public static int col_indentifier = 2;
        public static int col_keyword = col_indentifier + 1;
        public static int row_test_detail_start = 27;
        public static String regexKeywordString = @"(?i)Item";
        public static int row_offset_result_title = 1;                                // offset from the row of identifier regex "Item"
        public static int row_offset_bugstatus_title = row_offset_result_title;       // offset from the row of identifier regex "Item"
        public static int row_offset_buglist_title = row_offset_result_title + 1;     // offset from the row of identifier regex "Item"
        public static int col_offset_result_title = 1;                                // offset from the column of identifier regex "Item"
        public static int col_offset_bugstatus_title = col_offset_result_title + 2;   // offset from the column of identifier regex "Item"
        public static int col_offset_buglist_title = col_offset_result_title;         // offset from the column of identifier regex "Item"
        public static String regexResultString = @"^(?i)\s*Result\s*$";
        public static String regexBugStatusString = @"^(?i)\s*Bug Status\s*$";
        public static String regexBugListString = @"^(?i)\s*Bug List\s*$";

        public static String PASS_str = "Pass";
        public static String CONDITIONAL_PASS_str = "Conditional Pass";
        public static String FAIL_str = "Fail";
        public static String WAIVED_str = "Waived";

        public static Boolean Auto_Correct_Sheetname = false;

        public static KeywordReportHeader DefaultKeywordReportHeader = new KeywordReportHeader();

        public static int PassCnt_at_row = 21, PassCnt_at_col = 5;
        public static int FailCnt_at_row = 21, FailCnt_at_col = 7;
        //public static int TotalCnt_at_row = 21, TotalCnt_at_col = 9;
        public static int ConditionalPass_string_at_row = 21, ConditionalPass_string_at_col = 8;
        public static int ConditionalPassCnt_at_row = 21, ConditionalPassCnt_at_col = 9;

        private static List<TestPlanKeyword> global_keyword_list = new List<TestPlanKeyword>();
        private static Boolean global_keyword_available;
        public static List<TestPlanKeyword> GetGlobalKeywordList()
        {
            if (global_keyword_available)
            {
                return global_keyword_list;
            }
            else
            {
                return new List<TestPlanKeyword>();
            }
        }
        public static void ClearGlobalKeywordList()
        {
            global_keyword_available = false;
            global_keyword_list.Clear();
        }

        static public List<TestPlanKeyword> ListKeyword_SingleReport(TestPlan plan)
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
            RegexStringValidator result_keyword_Regex = new RegexStringValidator(regexResultString);
            RegexStringValidator bug_status_keyword_Regex = new RegexStringValidator(regexBugStatusString);
            RegexStringValidator bug_list_keyword_Regex = new RegexStringValidator(regexBugListString);
            for (int row_index = row_test_detail_start; row_index <= row_end; row_index++)
            {
                String identifier_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, col_indentifier),
                        result_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index + row_offset_result_title,
                                                                            col_indentifier + col_offset_result_title),
                        bug_status_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index + row_offset_bugstatus_title,
                                                                            col_indentifier + col_offset_bugstatus_title),
                        bug_list_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index + row_offset_buglist_title,
                                                                            col_indentifier + col_offset_buglist_title),
                        keyword_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, col_keyword);
                int regex_step = 0;
                try
                {
                    // Attempt validation.
                    // regex false (not a keyword row) then jumping to catch(); 
                    identifier_keyword_Regex.Validate(identifier_text); regex_step++;
                    // regex true, next step is to check the rest of field to validate
                    // 1. Check "Result" title
                    result_keyword_Regex.Validate(result_text); regex_step++;
                    bug_status_keyword_Regex.Validate(bug_status_text); regex_step++;
                    bug_list_keyword_Regex.Validate(bug_list_text); regex_step++;

                    if (keyword_text == "") { ConsoleWarning("Empty Keyword", row_index); continue; }
                    if (KeywordAtRow.ContainsKey(keyword_text)) { ConsoleWarning("Duplicated Keyword in single report", row_index); continue; }
                    KeywordAtRow.Add(keyword_text, row_index);
                }
                catch (ArgumentException e)
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
                            ConsoleWarning(regexBugStatusString, row_index + row_offset_result_title,
                                                                 col_indentifier + col_offset_result_title);
                            break;
                        case 2:
                            // Not a "Bug Status" 
                            ConsoleWarning(regexBugStatusString, row_index + row_index + row_offset_bugstatus_title,
                                                                 col_indentifier + col_offset_bugstatus_title);
                            break;
                        case 3:
                            // Not a "Bug List" 
                            ConsoleWarning(regexBugListString, row_index + row_offset_buglist_title,
                                                                 col_indentifier + col_offset_buglist_title);
                            break;
                        default:
                            break;
                    }
                    continue;
                }
            }

            List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
            foreach (String key in KeywordAtRow.Keys)
            {
                int row_keyword = KeywordAtRow[key];
                // col_keyword is currently fixed value
                int row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list;

                row_result = row_bug_status = row_keyword + 1;
                row_bug_list = row_keyword + 2;
                col_result = col_bug_list = col_keyword + 1;
                col_bug_status = col_keyword + 3;
                ret.Add(new TestPlanKeyword(key, plan.ExcelFile, plan.ExcelSheet, KeywordAtRow[key], col_keyword,
                    row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list));
            }
            return ret;
        }

        static public List<TestPlanKeyword> ListAllKeyword(List<TestPlan> DoPlan)
        {
            List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
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
                    List<TestPlanKeyword> plan_keyword = ListKeyword_SingleReport(plan);
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
                        ConsoleWarning("Test Plan null keyword list Error occurred:" + plan.ExcelSheet + "@" + plan.ExcelFile);
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
                        ConsoleWarning("Test Plan Unknown Error occurred:" + plan.ExcelSheet + "@" + plan.ExcelFile);
                    }
                    ret_not_report_log.Add(fail_log);
                    plan.CloseDetailExcel();
                }
            }
            ReportGenerator.excel_not_report_log.AddRange(ret_not_report_log);
            return ret;
        }

        static public List<TestPlanKeyword> FilterSingleReportKeyword(List<TestPlanKeyword> keyword_list, String workbook, String worksheet)
        {
            List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
            foreach (TestPlanKeyword kw in keyword_list)
            {
                if ((kw.Workbook == workbook) && (kw.Worksheet == worksheet))
                {
                    ret.Add(kw);
                }
            }
            return ret;
        }

        static public List<TestPlanKeyword> FilterSingleReportKeyword_check_only_worksheet(List<TestPlanKeyword> keyword_list, String worksheet)
        {
            List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
            foreach (TestPlanKeyword kw in keyword_list)
            {
                if ((kw.Worksheet == worksheet))
                {
                    ret.Add(kw);
                }
            }
            return ret;
        }

        static public Boolean IsAnyKeywordInReport(List<TestPlanKeyword> keyword_list, String workbook, String worksheet)
        {
            Boolean b_ret = false;
            foreach (TestPlanKeyword kw in keyword_list)
            {
                if ((kw.Workbook == workbook) && (kw.Worksheet == worksheet))
                {
                    b_ret = true;
                    break;
                }
            }
            return b_ret;
        }

        static public List<TestPlanKeyword> ListAllDuplicatedKeyword(List<TestPlanKeyword> keyword_list)
        {
            Dictionary<String, TestPlanKeyword> for_checking_duplicated = new Dictionary<String, TestPlanKeyword>();
            Dictionary<String, List<TestPlanKeyword>> dic_ret_kw_list = new Dictionary<String, List<TestPlanKeyword>>();

            foreach (TestPlanKeyword keyword in keyword_list)
            {
                String kw = keyword.Keyword;
                if (for_checking_duplicated.ContainsKey(kw))
                {
                    // found duplicated item
                    // first to check if already duplicated before
                    if (!dic_ret_kw_list.ContainsKey(kw))
                    {
                        // 1st time duplicated so that not available in dic_ret_kw_list
                        // then it is necessary to create a new item in dic_ret_kw_list
                        List<TestPlanKeyword> new_duplicated_list = new List<TestPlanKeyword>();
                        new_duplicated_list.Add(for_checking_duplicated[kw]);
                        dic_ret_kw_list.Add(kw, new_duplicated_list);
                    }
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

            List<TestPlanKeyword> ret_dup_kw_list = new List<TestPlanKeyword>();
            foreach (String kw in dic_ret_kw_list.Keys)
            {
                ret_dup_kw_list.AddRange(dic_ret_kw_list[kw]);
            }
            return ret_dup_kw_list;
        }

        static public List<String> ListDuplicatedKeywordString(List<TestPlanKeyword> keyword_list)
        {
            SortedSet<String> check_duplicated_keyword = new SortedSet<String>();
            List<TestPlanKeyword> duplicate_keyword_list = ListAllDuplicatedKeyword(keyword_list);
            foreach (TestPlanKeyword keyword in duplicate_keyword_list)
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

        static public void WriteBugCountOnKeywordReport(TestPlanKeyword keyword, Worksheet result_worksheet, IssueCount severity_count)
        {
            // Write severity count of all keywrod isseus
            List<StyleString> bug_status_string = new List<StyleString>();
            int issue_count;
            issue_count = severity_count.Severity_A;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "A", Issue.A_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0A", Color.Black));
            }
            //bug_status_string.Add(new StyleString(",", Color.Black));
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.Severity_B;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "B", Issue.B_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0B", Color.Black));
            }
            //bug_status_string.Add(new StyleString(",", Color.Black));
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 1, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.Severity_C;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "C", Issue.C_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0C", Color.Black));
            }
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 2, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.Severity_D;
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + "D", Issue.D_ISSUE_COLOR));
            }
            else
            {
                bug_status_string.Add(new StyleString("0D", Color.Black));
            }
            StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 3, bug_status_string);
            bug_status_string.Clear();

            issue_count = severity_count.TotalWaived();
            if (issue_count > 0)
            {
                bug_status_string.Add(new StyleString(issue_count.ToString() + " Waived", Issue.WAIVED_ISSUE_COLOR));
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

        static public void WriteKeywordConclusionOnKeywordReport(TestPlanKeyword keyword, Worksheet result_worksheet, IssueCount severity_count,
                                    out Boolean pass, out Boolean fail, out Boolean conditional_pass)
        {
            String pass_fail_str;
            pass = fail = conditional_pass = false;

            if (severity_count.NotClosedCount() == 0)
            {
                // all issue closed
                pass_fail_str = PASS_str;
                pass = true;
            }
            else if (severity_count.ABC_non_Wavied_IssueCount() > 0)
            {
                // any issue of ABC, non-closed & non-waived issue 
                pass_fail_str = FAIL_str;
                fail = true;
            }
            else
            {
                // only D or waived issue
                pass_fail_str = CONDITIONAL_PASS_str;
                conditional_pass = true;
            }
            ExcelAction.SetCellValue(result_worksheet, keyword.ResultListAtRow, keyword.ResultListAtColumn, pass_fail_str);
        }

        static public bool KeywordIssueGenerationTaskV4(List<String> file_list, String src_dir, String dest_dir = "")
        {
            // Clear keyword log report data-table
            ReportGenerator.excel_not_report_log.Clear();
            // 0.1 List all files under report_root_dir.
            // This is done outside and result is the input paramemter file_list
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
            List<TestPlanKeyword> keyword_list = ListAllKeyword(do_plan);
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
                        keyword.KeywordIssues.Add(issue);       // keep issue with keyword so that it can be used later.
                    }
                }
                keyword.IssueList = id_list;
                description_list = StyleString.ExtendIssueDescription(id_list, ReportGenerator.global_issue_description_list_severity);
                keyword.IssueDescriptionList = description_list;
            }

            // Output keyword list log excel here.
            String out_dir = (dest_dir != "") ? dest_dir : src_dir;
            KeyWordListReport.OutputKeywordLog(out_dir, keyword_list, ReportGenerator.excel_not_report_log, output_keyword_issue: true);
            //KeyWordListReport.OutputKeywordLog(out_dir, keyword_list, ReportGenerator.excel_not_report_log, output_keyword_issue: false);

            // Load global_keyword_list here
            global_keyword_list = keyword_list;
            global_keyword_available = true;

            //
            // 3. Go throught each report excel and generate keyword report for each one.
            //
            foreach (TestPlan plan in do_plan)
            {
                String full_filename = plan.ExcelFile;
                String sheet_name = plan.ExcelSheet;

                // 3.0 if there isn't any keyword in this plan, just continue to next plan
                //     
                List<TestPlanKeyword> ws_keyword_list = FilterSingleReportKeyword(keyword_list, full_filename, sheet_name);
                if (ws_keyword_list.Count <= 0)
                {
                    continue;
                }

                // 3.1. Open Excel and find the sheet
                // File exist check is done outside
                Workbook wb_keyword_issue = ExcelAction.OpenExcelWorkbook(full_filename);
                if (wb_keyword_issue == null)
                {
                    ConsoleWarning("ERR: Open workbook in V4: " + full_filename);
                    continue;
                }

                // 3.2 Open worksheet
                Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_keyword_issue, sheet_name);
                if (result_worksheet == null)
                {
                    ConsoleWarning("ERR: Open worksheet in V4: " + full_filename + " sheet: " + sheet_name);
                    continue;
                }

                //
                // 3.3. input:  IssueDescriptionList of Keyword
                //    output: write color_description_list 
                //         

                // 3.3.2 Write keyword-related formatted issue descriptions 
                //       also count how many "Pass" or how many "Fail"
                int pass_count = 0, fail_count = 0, conditional_pass_count = 0;
                //foreach (TestPlanKeyword keyword in keyword_list)
                foreach (TestPlanKeyword keyword in ws_keyword_list)
                {
                    // Only write to keyword on currently open sheet
                    //if (keyword.Worksheet == sheet_name)
                    {
                        // write issue description list
                        StyleString.WriteStyleString(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn, keyword.IssueDescriptionList);

                        // write issue count of each severity
                        IssueCount severity_count = keyword.Calculate_Issue();
                        WriteBugCountOnKeywordReport(keyword, result_worksheet, severity_count);

                        // write conclusion of each keyword
                        Boolean pass, fail, conditional_pass;
                        WriteKeywordConclusionOnKeywordReport(keyword, result_worksheet, severity_count, out pass, out fail, out conditional_pass);
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
                        ExcelAction.AutoFit_Row(result_worksheet, keyword.ResultListAtRow);
                        ExcelAction.AutoFit_Row(result_worksheet, keyword.BugListAtRow);
                        // issue_count = severity_count.Severity_A + severity_count.Severity_B + severity_count.Severity_C;
                        //if (issue_count >= 1)
                        int issue_count = severity_count.NotClosedCount();
                        if (issue_count > 0)
                        {
                            double single_row_height = ExcelAction.Get_Row_Height(result_worksheet, keyword.BugListAtRow);
                            double new_row_height = single_row_height * issue_count * 0.8 + 0.2;
                            ExcelAction.Set_Row_Height(result_worksheet, keyword.BugListAtRow, new_row_height);
                        }
                        else
                        {
                            // Hide bug list row only when there isn't any non-closed issue at all (all issues must be closed)
                            ExcelAction.Hide_Row(result_worksheet, keyword.BugListAtRow);
                        }
                        //ExcelAction.CellTextAlignLeft(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn);
                        ExcelAction.CellTextAlignUpperLeft(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn);
                    }
                }

                // 3.3.3 Update Conclusion
                //const int PassCnt_at_row = 21, PassCnt_at_col = 5;
                //const int FailCnt_at_row = 21, FailCnt_at_col = 7;
                //const int TotalCnt_at_row = 21, TotalCnt_at_col = 9;
                //const int Judgement_at_row = 9, Judgement_at_col = 4;
                String judgement_str;
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
                ExcelAction.SetCellValue(result_worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col, judgement_str);
                // always update Test End Period to today
                String end_date = DateTime.Now.ToString("yyyy/MM/dd");
                ExcelAction.SetCellValue(result_worksheet, KeywordReportHeader.Test_Period_End_at_row, KeywordReportHeader.Test_Period_End_at_col, end_date);
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
            }

            // Output updated report with recommended sheetname.
            if (KeywordReport.Auto_Correct_Sheetname == true)
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
                            ConsoleWarning("ERR: Open workbook in auto-correct-worksheet-name: " + full_filename);
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

        static public Boolean UpdateKeywordReportHeader_full(Worksheet report_worksheet, KeywordReportHeader header)
        {
            Boolean b_ret = false;
            try
            {
                //@"Update_Report_Title_by_Sheetname",        @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Report_Title_by_Sheetname)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Title_at_row, KeywordReportHeader.Title_at_col, header.Report_Title);
                }

                //@"Update_Model_Name",                       @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Model_Name)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Model_Name_at_row, KeywordReportHeader.Model_Name_at_col, header.Model_Name);
                }

                //@"Update_Part_No",                          @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Part_No)
                {
                    String output_part_no = header.Part_No + "-" + header.Report_SheetName;
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Part_No_at_row, KeywordReportHeader.Part_No_at_col, output_part_no);
                }

                //@"Update_Panel_Module",                     @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Panel_Module)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Panel_Module_at_row, KeywordReportHeader.Panel_Module_at_col, header.Panel_Module);
                }

                //@"Update_TCON_Board",                       @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_TCON_Board)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.TCON_Board_at_row, KeywordReportHeader.TCON_Board_at_col, header.TCON_Board);
                }

                //@"Update_AD_Board",                         @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_AD_Board)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.AD_Board_at_row, KeywordReportHeader.AD_Board_at_col, header.AD_Board);
                }

                //@"Update_Power_Board",                      @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Power_Board)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Power_Board_at_row, KeywordReportHeader.Power_Board_at_col, header.Power_Board);
                }

                //@"Update_Smart_BD_OS_Version",              @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Smart_BD_OS_Version)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Smart_BD_OS_Version_at_row, KeywordReportHeader.Smart_BD_OS_Version_at_col, header.Smart_BD_OS_Version);
                }

                //@"Update_Touch_Sensor",                     @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Touch_Sensor)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Touch_Sensor_at_row, KeywordReportHeader.Touch_Sensor_at_col, header.Touch_Sensor);
                }

                //@"Update_Speaker_AQ_Version",               @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Speaker_AQ_Version)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Speaker_AQ_Version_at_row, KeywordReportHeader.Speaker_AQ_Version_at_col, header.Speaker_AQ_Version);
                }

                //@"Update_SW_PQ_Version",                    @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_SW_PQ_Version)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.SW_PQ_Version_at_row, KeywordReportHeader.SW_PQ_Version_at_col, header.SW_PQ_Version);
                }

                //@"Update_Test_Stage",                       @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Test_Stage)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_Stage_at_row, KeywordReportHeader.Test_Stage_at_col, header.Test_Stage);
                }

                //@"Update_Test_QTY_SN",                      @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Test_QTY_SN)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_QTY_SN_at_row, KeywordReportHeader.Test_QTY_SN_at_col, header.Test_QTY_SN);
                }

                //@"Update_Test_Period_Begin",                @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Test_Period_Begin)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_Period_Begin_at_row, KeywordReportHeader.Test_Period_Begin_at_col, header.Test_Period_Begin);
                }

                //@"Update_Test_Period_End",                  @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Test_Period_End)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Test_Period_End_at_row, KeywordReportHeader.Test_Period_End_at_col, header.Test_Period_End);
                }

                //@"Update_Judgement",                        @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Judgement)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col, header.Judgement);
                }

                //@"Update_Tested_by",                        @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Tested_by)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Tested_by_at_row, KeywordReportHeader.Tested_by_at_col, header.Tested_by);
                }

                //@"Update_Approved_by",                      @"true",
                if (KeywordReport.DefaultKeywordReportHeader.Update_Approved_by)
                {
                    ExcelAction.SetCellValue(report_worksheet, KeywordReportHeader.Approved_by_at_row, KeywordReportHeader.Approved_by_at_col, header.Approved_by);
                }

                b_ret = true;
            }
            catch (Exception ex)
            {
            }

            return b_ret;
        }

        static public Boolean UpdateReportHeader(Worksheet ws, String Title = null, String SW_Version = null, String Test_Start = null, String Test_End = null,
                        String Judgement = null, String Template = null)
        {
            Boolean b_ret = false;
            // to-be-finished.
            if (Template != null)
            {

            }
            else
            {
                if (Title != null)
                {
                    ExcelAction.SetCellValue(ws, KeywordReportHeader.Title_at_row, KeywordReportHeader.Title_at_col, Title);
                }
                if (SW_Version != null)
                {
                    ExcelAction.SetCellValue(ws, KeywordReportHeader.SW_PQ_Version_at_row, KeywordReportHeader.SW_PQ_Version_at_col, Judgement);
                }
                if (Test_Start != null)
                {
                    ExcelAction.SetCellValue(ws, KeywordReportHeader.Test_Period_Begin_at_row, KeywordReportHeader.Test_Period_Begin_at_col, Test_Start);
                }
                if (Test_End != null)
                {
                    ExcelAction.SetCellValue(ws, KeywordReportHeader.Test_Period_End_at_row, KeywordReportHeader.Test_Period_End_at_col, Test_End);
                }
                if (Judgement != null)
                {
                    ExcelAction.SetCellValue(ws, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col, Judgement);
                }
            }
            b_ret = true;
            return b_ret;
        }

        public static Boolean UpdateAllHeader(List<String> report_list, String Title = null, String SW_Version = null, String Test_Start = null, String Test_End = null,
                                        String Judgement = null, String Template = null)
        {
            // Create a temporary test plan to includes all files listed in List<String> report_filename
            List<TestPlan> do_plan = TestPlan.CreateTempPlanFromFileList(report_list);

            foreach (TestPlan plan in do_plan)
            {
                String path = Storage.GetDirectoryName(plan.ExcelFile);
                String filename = Storage.GetFileName(plan.ExcelFile);
                String sheet_name = plan.ExcelSheet;
                TestPlan.ExcelStatus test_plan_status;

                test_plan_status = plan.OpenDetailExcel(ReadOnly: false);
                if (test_plan_status == TestPlan.ExcelStatus.OK)
                {
                    UpdateReportHeader(plan.TestPlanWorksheet, Title: Title, SW_Version: SW_Version, Test_Start: Test_Start,
                                            Test_End: Test_End, Judgement: Judgement, Template: Template);
                    plan.SaveDetailExcel(plan.ExcelFile);
                    plan.CloseDetailExcel();
                }
            }
            return true;
        }

        // This function is used to get judgement result (only read and no update to report) of keyword report
        static public Boolean GetJudgementValue(String report_workbook, String report_worksheet, out String judgement_str)
        {
            Boolean b_ret = false;
            String ret_str = ""; // default returning judgetment_str;

            // 1. Open Excel and find the sheet
            // File exist check is done outside
            Workbook wb_judgement = ExcelAction.OpenExcelWorkbook(report_workbook);
            if (wb_judgement == null)
            {
                ConsoleWarning("ERR: Open workbook in GetJudgementValue: " + report_workbook);
                judgement_str = ret_str;
                b_ret = false;
            }
            else
            {
                // 2 Open worksheet
                Worksheet ws_judgement = ExcelAction.Find_Worksheet(wb_judgement, report_worksheet);
                if (ws_judgement == null)
                {
                    ConsoleWarning("ERR: Open worksheet in GetJudgementValue: " + report_workbook + " sheet: " + report_worksheet);
                    judgement_str = ret_str;
                    b_ret = false;
                }
                else
                {
                    // 3. Get Judgement value
                    Object obj = ExcelAction.GetCellValue(ws_judgement, KeywordReportHeader.Judgement_at_row, KeywordReportHeader.Judgement_at_col);
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
        static public List<TestPlanKeyword> ListAllDetailedTestPlanKeywordTask(String report_root_dir, String output_filename)
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
            List<TestPlanKeyword> keyword_list = ListAllKeyword(do_plan);

            // Output keyword list log excel here.
            KeyWordListReport.OutputKeywordLog(report_root_dir, keyword_list, ReportGenerator.excel_not_report_log, output_filename);


            // Output updated report with recommended sheetname.
            if (KeywordReport.Auto_Correct_Sheetname == true)
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
                            ConsoleWarning("ERR: Open workbook in auto-correct-worksheet-name of ListAllDetailedTestPlanKeywordTask(): " + full_filename);
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

    }
}
