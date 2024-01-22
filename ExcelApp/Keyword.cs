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

    public class ReportKeyword
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

        public ReportKeyword() { TestPlanKeywordInit(); }
        public ReportKeyword(String Keyword, String Workbook = "", String Worksheet = "", int AtRow = 0, int AtColumn = 0,
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

        public static String regexKeywordString = @"(?i)Item";
        public static int col_keyword = TestReport.col_indentifier + 1;
        public static int row_offset_result_title = 1;                                // offset from the row of identifier regex "Item"
        public static int col_offset_result_title = 1;                                // offset from the column of identifier regex "Item"
        public static int row_offset_bugstatus_title = row_offset_result_title;       // offset from the row of identifier regex "Item"
        public static int col_offset_bugstatus_title = col_offset_result_title + 2;   // offset from the column of identifier regex "Item"
        public static int row_offset_buglist_title = row_offset_result_title + 1;     // offset from the row of identifier regex "Item"
        public static int col_offset_buglist_title = col_offset_result_title;         // offset from the column of identifier regex "Item"
        public static String regexResultString = @"^(?i)\s*Result\s*$";
        public static String regexBugStatusString = @"^(?i)\s*Bug Status\s*$";
        public static String regexBugListString = @"^(?i)\s*Bug List\s*$";

        static private Boolean CheckIfStringMeetsKeywordResultCondition(String text_to_check)
        {
            String regex = regexResultString;
            return TestReport.CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsKeywordBugStatusCondition(String text_to_check)
        {
            String regex = regexBugStatusString;
            return TestReport.CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        static private Boolean CheckIfStringMeetsKeywordBugListCondition(String text_to_check)
        {
            String regex = regexBugListString;
            return TestReport.CheckIfStringMeetsRegexString(text_to_check, regex);
        }

        // Visit report content and find out all keywords
        static public List<ReportKeyword> ListKeyword_SingleReport(TestPlan plan)
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
            for (int row_index = TestReport.row_test_detail_start; row_index <= row_end; row_index++)
            {
                String identifier_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, TestReport.col_indentifier),
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
                                                                 TestReport.col_indentifier + col_offset_result_title);
                            break;
                        case 2:
                            // Not a "Bug Status" 
                            LogMessage.CheckFunctionAtRowColumn(regexBugStatusString, row_index + row_index + row_offset_bugstatus_title,
                                                                 TestReport.col_indentifier + col_offset_bugstatus_title);
                            break;
                        case 3:
                            // Not a "Bug List" 
                            LogMessage.CheckFunctionAtRowColumn(regexBugListString, row_index + row_offset_buglist_title,
                                                                 TestReport.col_indentifier + col_offset_buglist_title);
                            break;
                        default:
                            break;
                    }
                    continue;
                }
            }

            List<ReportKeyword> ret = new List<ReportKeyword>();
            foreach (String key in KeywordAtRow.Keys)
            {
                int row_keyword = KeywordAtRow[key];
                // col_keyword is currently fixed value
                int row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list;

                row_result = row_bug_status = row_keyword + 1;
                row_bug_list = row_keyword + 2;
                col_result = col_bug_list = col_keyword + 1;
                col_bug_status = col_keyword + 3;
                ret.Add(new ReportKeyword(key, plan.ExcelFile, plan.ExcelSheet, KeywordAtRow[key], col_keyword,
                    row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list));
            }
            return ret;
        }

        static public List<ReportKeyword> ListAllDuplicatedKeyword(List<ReportKeyword> keyword_list)
        {
            Dictionary<String, ReportKeyword> for_checking_duplicated = new Dictionary<String, ReportKeyword>();
            Dictionary<String, List<ReportKeyword>> dic_ret_kw_list = new Dictionary<String, List<ReportKeyword>>();

            foreach (ReportKeyword keyword in keyword_list)
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
                        List<ReportKeyword> new_duplicated_list = new List<ReportKeyword>();
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

            List<ReportKeyword> ret_dup_kw_list = new List<ReportKeyword>();
            foreach (String kw in dic_ret_kw_list.Keys)
            {
                ret_dup_kw_list.AddRange(dic_ret_kw_list[kw]);
            }
            return ret_dup_kw_list;
        }

        // List all duplicated keywords based on complete keyword-list
        static public List<String> ListDuplicatedKeywordString(List<ReportKeyword> keyword_list)
        {
            SortedSet<String> check_duplicated_keyword = new SortedSet<String>();
            List<ReportKeyword> duplicate_keyword_list = ListAllDuplicatedKeyword(keyword_list);
            foreach (ReportKeyword keyword in duplicate_keyword_list)
            {
                check_duplicated_keyword.Add(keyword.Keyword);
            }
            List<String> ret_str_list = new List<String>();
            ret_str_list.AddRange(check_duplicated_keyword);
            return ret_str_list;
        }

        static public Boolean HideKeywordResultBugRow(String excel_filename, Worksheet ws)
        {
            // Clear Keyword bug result and hide 2 rows (only Item row left)
            Boolean b_ret = false;
            try
            {
                TestPlan tp = TestPlan.CreateTempPlanFromFile(excel_filename);
                tp.TestPlanWorksheet = ws;
                tp.ExcelSheet = ws.Name;
                List<ReportKeyword> keyword_list = ListKeyword_SingleReport(tp);
                foreach (ReportKeyword keyword in keyword_list)
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
                List<ReportKeyword> keyword_list = ListKeyword_SingleReport(tp);
                foreach (ReportKeyword keyword in keyword_list)
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

    }

}
