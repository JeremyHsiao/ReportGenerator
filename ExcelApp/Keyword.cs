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

        public IssueCount Calculate_Issue_GE_Stage_1_0()
        {
            IssueCount ret_ic = new IssueCount();
            foreach (Issue issue in this.KeywordIssues)
            {
                if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
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

        public static int PassCnt_at_row = 21, PassCnt_at_col = 5;
        public static int FailCnt_at_row = 21, FailCnt_at_col = 7;
        public static int TotalCnt_at_row = 21, TotalCnt_at_col = 9;
        public static int Judgement_at_row = 9, Judgement_at_col = 4;
        public static int Judgement_string_at_row = 9, Judgement_string_at_col = 2;

        static public List<TestPlanKeyword> ListKeyword(TestPlan plan)
        {
            //
            // 2. Find out Printable Area
            //
            // Assummed that Printable area always starting at $A$1 (also data processing area)
            // So excel data processing area ends at Printable area (row_count,col_count)
            Worksheet ws_testplan = plan.TestPlanWorksheet;
            Range rngPrintable = ExcelAction.GetWorksheetPrintableRange(ws_testplan);
            int row_print_area = ExcelAction.Get_Range_RowNumber(rngPrintable);
            int column_print_area = ExcelAction.Get_Range_ColumnNumber(rngPrintable);

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
            for (int row_index = row_test_detail_start; row_index <= row_print_area; row_index++)
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
                    if (KeywordAtRow.ContainsKey(keyword_text)) { ConsoleWarning("Duplicated Keyword", row_index); continue; }
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
            List<NotReportFileRecord> ret_not_report_log = new List<NotReportFileRecord>();

            foreach (TestPlan plan in DoPlan)
            {
                String path = Storage.GetDirectoryName(plan.ExcelFile);
                String filename = Storage.GetFileName(plan.ExcelFile);
                NotReportFileRecord fail_log = new NotReportFileRecord(path, filename);

                TestPlan.ExcelStatus test_plan_status;
                test_plan_status = plan.OpenDetailExcel();
                if (test_plan_status == TestPlan.ExcelStatus.OK)
                {
                    List<TestPlanKeyword> plan_keyword = ListKeyword(plan);
                    plan.CloseDetailExcel();
                    if (plan_keyword != null)
                    {
                        if (plan_keyword.Count() > 0)
                        {
                            ret.AddRange(plan_keyword);
                            fail_log.SetFlagOK(openfileOK: true, findWorksheetOK: true, findAnyKeyword: true);
                            // not adding ok report log at the moment
                            //ret_not_report_log.Add(fail_log);
                        }
                        else
                        {
                            fail_log.SetFlagOK(openfileOK: true, findWorksheetOK: true);
                            fail_log.SetFlagFail(findNoKeyword: true);
                            ret_not_report_log.Add(fail_log);
                        }
                    }
                    else // (null) 
                    {
                        fail_log.SetFlagOK(openfileOK: true, findWorksheetOK: true);
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
                        fail_log.SetFlagOK(openfileOK: true); 
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

        static public bool KeywordIssueGenerationTaskV4(List<String> report_filename, String src_dir, String dest_dir = "")
        {
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
            ReportGenerator.excel_not_report_log.Clear();
            List<TestPlanKeyword> keyword_list = ListAllKeyword(do_plan);
            
            // Output keyword list log excel here.
            String out_dir = (dest_dir!="")?dest_dir:src_dir;
            KeyWordListReport.OutputKeywordLog(out_dir, keyword_list, ReportGenerator.excel_not_report_log);

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
                int pass_count = 0, fail_count = 0;
                //foreach (TestPlanKeyword keyword in keyword_list)
                foreach (TestPlanKeyword keyword in ws_keyword_list)
                {
                    // Only write to keyword on currently open sheet
                    //if (keyword.Worksheet == sheet_name)
                    {
                        // write issue description list
                        StyleString.WriteStyleString(result_worksheet, keyword.BugListAtRow, keyword.BugListAtColumn, keyword.IssueDescriptionList);

                        // Write severity count of all keywrod isseus
                        IssueCount severity_count = keyword.Calculate_Issue_GE_Stage_1_0();
                        List<StyleString> bug_status_string = new List<StyleString>();
                        int issue_count;
                        issue_count = severity_count.Severity_A;
                        if (issue_count>0)
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
                        if (issue_count>0)
                        {
                            bug_status_string.Add(new StyleString(issue_count.ToString() + "B", Issue.B_ISSUE_COLOR));
                        }
                        else
                        {
                            bug_status_string.Add(new StyleString("0B", Color.Black));
                        }
                        //bug_status_string.Add(new StyleString(",", Color.Black));
                        StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn+1, bug_status_string);
                        bug_status_string.Clear();

                        issue_count = severity_count.Severity_C;
                        if (issue_count>0)
                        {
                            bug_status_string.Add(new StyleString(issue_count.ToString() + "C", Issue.C_ISSUE_COLOR));
                        }
                        else
                        {
                            bug_status_string.Add(new StyleString("0C", Color.Black));
                        }
                        StyleString.WriteStyleString(result_worksheet, keyword.BugStatusAtRow, keyword.BugStatusAtColumn + 2, bug_status_string);
                        bug_status_string.Clear();

                        // Output Result:
                        //// >0A: NG
                        //// >0B: Defect found
                        //// >0C: Minor issue only
                        //// 0A0B0C: Good
                        // Pass: no issue
                        // Fail: any issue
                        // (A/B/C) = (xx/oo/vv)
                        //List<StyleString> result_string = new List<StyleString>();
                        String pass_fail_str = "Fail";
                        if (severity_count.Severity_A > 0)
                        {
                            //String Has_A_Issue_str = "Fail"; // "NG"
                            //Color Has_A_Issue_color = Issue.A_ISSUE_COLOR;
                            //result_string.Add(new StyleString(Has_A_Issue_str, Has_A_Issue_color));
                            pass_fail_str = "Fail";
                            fail_count++;
                        }
                        else if (severity_count.Severity_B > 0)
                        {
                            //String No_A_Has_B_Issue_str = "Fail"; // "Defect found"
                            //Color No_A_Has_B_Issue_color = Issue.A_ISSUE_COLOR;
                            //result_string.Add(new StyleString(No_A_Has_B_Issue_str, No_A_Has_B_Issue_color));
                            pass_fail_str = "Fail";
                            fail_count++;
                        }
                        else if (severity_count.Severity_C > 0)
                        {
                            //String No_AB_Has_C_Issue_str = "Fail"; // "Minor Issue only"
                            //Color No_AB_Has_C_Issue_color = Issue.A_ISSUE_COLOR;
                            //result_string.Add(new StyleString(No_AB_Has_C_Issue_str, No_AB_Has_C_Issue_color));
                            pass_fail_str = "Fail";
                            fail_count++;
                        }
                        else 
                        {
                            //String No_Issue_str = "Pass"; 
                            //Color No_Issue_color = Color.Lime;
                            //result_string.Add(new StyleString(No_Issue_str, No_Issue_color));
                            pass_fail_str = "Pass";
                            pass_count++;
                        }
                        //StyleString.WriteStyleString(result_worksheet, keyword.ResultListAtRow, keyword.ResultListAtColumn, result_string);
                        // Pass/Fail only value is updated.
                        ExcelAction.SetCellValue(result_worksheet, keyword.ResultListAtRow, keyword.ResultListAtColumn, pass_fail_str);

                        ExcelAction.AutoFit_Row(result_worksheet, keyword.ResultListAtRow);
                        ExcelAction.AutoFit_Row(result_worksheet, keyword.BugListAtRow);
                        issue_count = severity_count.Severity_A + severity_count.Severity_B + severity_count.Severity_C;
                        if (issue_count >= 1)
                        {
                            double single_row_height = ExcelAction.Get_Row_Height(result_worksheet, keyword.BugListAtRow);
                            ExcelAction.Set_Row_Height(result_worksheet, keyword.BugListAtRow, single_row_height * issue_count * 0.8 + 0.2);
                        }
                        else
                        {
                            // hide row if no bug list at all
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
                    judgement_str = "Fail";
                }
                else
                {
                    // pass
                    judgement_str = "Pass";
                }
                ExcelAction.SetCellValue(result_worksheet, PassCnt_at_row, PassCnt_at_col, pass_count);
                ExcelAction.SetCellValue(result_worksheet, FailCnt_at_row, FailCnt_at_col, fail_count);
                ExcelAction.SetCellValue(result_worksheet, TotalCnt_at_row, TotalCnt_at_col, fail_count + pass_count);
                ExcelAction.SetCellValue(result_worksheet, Judgement_at_row, Judgement_at_col, judgement_str);

                // 3.4. Save the file to either 
                //  (1) filename with yyyyMMddHHmmss if dest_dir is not specified
                //  (2) the same filename but to the sub-folder of same strucure under new root-folder "dest_dir"
                String dest_filename = DecideDestinationFilename(src_dir, dest_dir, full_filename);
                String dest_filename_dir = Storage.GetDirectoryName(dest_filename);
                // if parent directory does not exist, create recursively all parents
                Storage.CreateDirectory(dest_filename_dir, auto_parent_dir: true);
                ExcelAction.CloseExcelWorkbook(wb_keyword_issue, SaveChanges: true, AsFilename: dest_filename);
            }

            //// Output updated report with recommended sheetname.
            //if (KeywordReport.Auto_Correct_Sheetname == true)
            //{
            //    // ReportGenerator.excel_not_report_log
            //    foreach (NotReportFileRecord item in ReportGenerator.excel_not_report_log)
            //    {
            //        Boolean openfileOK, findWorksheetOK, findAnyKeyword, otherFailure;
            //        item.GetFlagValue(out openfileOK, out findWorksheetOK, out findAnyKeyword, out otherFailure);
            //        if((openfileOK==true)&&(findWorksheetOK==false)&&(otherFailure==false))
            //        {
            //            String full_filename, sheet_name;
            //            // Open Excel and find the sheet
            //            full_filename = item.GetFullFilePath();
            //            Workbook wb = ExcelAction.OpenExcelWorkbook(full_filename);
            //            if (wb == null)
            //            {
            //                ConsoleWarning("ERR: Open workbook in V4_2: " + full_filename);
            //                continue;
            //            }

            //            // Use active worksheet and rename it.
            //            Worksheet ws = ExcelAction.Find_Worksheet(wb, sheet_name);
            //            if (result_worksheet == null)
            //            {
            //                ConsoleWarning("ERR: Open worksheet in V4_2: " + full_filename + " sheet: " + sheet_name);
            //                continue;
            //            }
            //        }
            //  }
            //}

            return true;
        }

        static private String DecideDestinationFilename(String src_dir, String dest_dir, String full_filename)
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
        static public List<TestPlanKeyword> ListAllDetailedTestPlanKeywordTask(String filename, String report_root_dir)
        {
            // Full file name exist checked before executing task

            List<TestPlanKeyword> keyword_list = new List<TestPlanKeyword>();

            // read test-plan sheet NG and return if NG
            List<TestPlan> testplan = TestReport.ReadTestPlanFromStandardTestReport(filename);
            if (testplan == null) { return keyword_list; }

            // all input parameters has been checked successfully, so generate
            List<TestPlan> do_plan = TestPlan.ListDoPlan(testplan);
            foreach (TestPlan plan in do_plan)
            {
                plan.ExcelFile = report_root_dir + @"\" + plan.ExcelFile;
            }
            ReportGenerator.excel_not_report_log.Clear();
            keyword_list = ListAllKeyword(do_plan);

            return keyword_list;
        }

    }
}
