using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace ExcelReportApplication
{
    public class TestPlan
    {
        private String group;
        private String summary;
        private String assignee;
        private String do_or_not;
        private String category;
        private String subpart;

        // The following members will be used but not part of the test plan in Standard Test Report. (out-of-band data)
        private String from;
        private String path;
        private String sheet;
        private Workbook wb_testplan;
        private Worksheet ws_testplan;

        public String Group   // property
        {
            get { return group; }   // get method
            set { group = value; }  // set method
        }

        public String Summary   // property
        {
            get { return summary; }   // get method
            set { summary = value; }  // set method
        }

        public String Assignee   // property
        {
            get { return assignee; }   // get method
            set { assignee = value; }  // set method
        }

        public String DoOrNot   // property
        {
            get { return do_or_not; }   // get method
            set { do_or_not = value; }  // set method
        }

        public String Category   // property
        {
            get { return category; }   // get method
            set { category = value; }  // set method
        }

        public String Subpart   // property
        {
            get { return subpart; }   // get method
            set { subpart = value; }  // set method
        }

        public String BackupSource   // property
        {
            get { return from; }   // get method
            set { from = value; }  // set method
        }

        public String ExcelSheet   // property
        {
            get { return sheet; }   // get method
            set { sheet = value; }  // set method
        }

        public String ExcelFile   // property
        {
            get { return path; }   // get method
            set { path = value; }  // set method
        }

        public TestPlan()
        {
        }

        public TestPlan(List<String> members)
        {
            this.group = members[(int)TestPlanMemberIndex.GROUP];
            this.summary = members[(int)TestPlanMemberIndex.SUMMARY];
            this.assignee = members[(int)TestPlanMemberIndex.ASSIGNEE];
            this.do_or_not = members[(int)TestPlanMemberIndex.DO_OR_NOT];
            this.category = members[(int)TestPlanMemberIndex.CATEGORY];
            this.subpart = members[(int)TestPlanMemberIndex.SUBPART];
        }

        public enum TestPlanMemberIndex
        {
            GROUP = 0,
            SUMMARY,
            ASSIGNEE,
            DO_OR_NOT,
            CATEGORY,
            SUBPART,
        }

        public static int TestPlanMemberCount = Enum.GetNames(typeof(TestPlanMemberIndex)).Length;

        public const string col_Group = "Test Group";
        public const string col_Summary = "Summary";
        public const string col_Assignee = "Assignee";
        public const string col_DoOrNot = "Do or Not";
        public const string col_Category = "Test Case Category";
        public const string col_Subpart = "Subpart";
        // The sequence of this String[] must be aligned with enum TestPlanMemberIndex
        static public String[] TestPlanMemberColumnName = { col_Group, col_Summary, col_Assignee, col_DoOrNot, col_Category, col_Subpart };

        public static int NameDefinitionRow_TestPlan = 2;
        public static int DataBeginRow_TestPlan = 3;

        public static List<TestPlan> ListDoPlan(List<TestPlan> testplan)
        {
            List<TestPlan> do_plan = new List<TestPlan>();
            foreach (TestPlan tp in testplan)
            {
                if (tp.DoOrNot == "V")
                {
                    do_plan.Add(tp);
                }
            }
            return do_plan;
        }

        public static List<TestPlan> LoadTestPlanSheet(Worksheet testplan_ws)
        {
            List<TestPlan> ret_testplan = new List<TestPlan>();

            // Create index for each column name
            Dictionary<string, int> col_name_list = ExcelAction.CreateTableColumnIndex(testplan_ws, NameDefinitionRow_TestPlan);

            // Get the last (row,col) of excel
            Range rngLast = ExcelAction.GetWorksheetAllRange(testplan_ws);
            int row_end = rngLast.Row;
            // Visit all rows and add content 
            for (int index = DataBeginRow_TestPlan; index <= row_end; index++)
            {
                List<String> members = new List<String>();
                for (int member_index = 0; member_index < TestPlanMemberCount; member_index++)
                {
                    int col_index = col_name_list[TestPlanMemberColumnName[member_index]];
                    String str = ExcelAction.GetCellTrimmedString(testplan_ws, index, col_index);
                    if (str == "")
                    {
                        break; // cannot be empty value; skip to next row
                    }
                    members.Add(str);
                }
                if (members.Count == TestPlanMemberCount)
                {
                    TestPlan tp = new TestPlan(members);
                    ret_testplan.Add(tp);
                }
            }
            return ret_testplan;
        }

        public enum ExcelStatus
        {
            OK = 0,
            INIT_STATE,
            ERR_OpenDetailExcel_OpenExcelWorkbook,
            ERR_OpenDetailExcel_Find_Worksheet,
            ERR_CloseDetailExcel_wb_null,
            ERR_SaveChangesAndCloseDetailExcel_wb_null,
            ERR_NOT_DEFINED,
            EX_OpenDetailExcel,
            EX_CloseDetailExcel,
            EX_SaveChangesAndCloseDetailExcel,
            MAX_NO
        };

        public ExcelStatus OpenDetailExcel()
        {
            try
            {
                Workbook wb;

                // Open excel (read-only & corrupt-load)
                wb = ExcelAction.OpenExcelWorkbook(path);

                if (wb == null)
                {
                    return ExcelStatus.ERR_OpenDetailExcel_OpenExcelWorkbook;
                }

                Worksheet ws = ExcelAction.Find_Worksheet(wb, sheet);
                if (ws == null)
                {
                    return ExcelStatus.ERR_OpenDetailExcel_Find_Worksheet;
                }
                else
                {
                    wb_testplan = wb;
                    ws_testplan = ws;
                    return ExcelStatus.OK;
                }
            }
            catch
            {
                return ExcelStatus.EX_OpenDetailExcel;
            }
            // Not needed because never reaching here
            //return ExcelStatus.ERR_NOT_DEFINED;
        }

        public ExcelStatus CloseIssueListExcel()
        {
            try
            {
                if (wb_testplan == null)
                {
                    return ExcelStatus.ERR_CloseDetailExcel_wb_null;
                }
                ExcelAction.CloseExcelWorkbook(wb_testplan, SaveChanges: false);
                ws_testplan = null;
                wb_testplan = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_testplan = null;
                wb_testplan = null;
                return ExcelStatus.EX_CloseDetailExcel;
            }
        }

        public ExcelStatus SaveChangesAndCloseIssueListExcel(String dest_filename)
        {
            try
            {
                if (wb_testplan == null)
                {
                    return ExcelStatus.ERR_SaveChangesAndCloseDetailExcel_wb_null;
                }
                ExcelAction.CloseExcelWorkbook(wb_testplan, SaveChanges: true, AsFilename: dest_filename);
                ws_testplan = null;
                wb_testplan = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_testplan = null;
                wb_testplan = null;
                return ExcelStatus.EX_SaveChangesAndCloseDetailExcel;
            }
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

        public List<TestPlanKeyword> ListKeyword()
        {
            //
            // 2. Find out Printable Area
            //
            // Assummed that Printable area always starting at $A$1 (also data processing area)
            // So excel data processing area ends at Printable area (row_count,col_count)
            Range rngPrintable = ExcelAction.GetWorksheetPrintableRange(ws_testplan);
            int row_print_area = rngPrintable.Rows.Count;
            int column_print_area = rngPrintable.Columns.Count;

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
                String  identifier_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, col_indentifier),
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
                            // Not a keyword identifier 
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

            List<TestPlanKeyword> ret =  new List<TestPlanKeyword> ();
            foreach (String key in KeywordAtRow.Keys)
            {
                int row_keyword = KeywordAtRow[key];
                // col_keyword is currently fixed value
                int row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list;

                row_result = row_bug_status = row_keyword + 1;
                row_bug_list = row_keyword + 2;
                col_result = col_bug_list = col_keyword + 1;
                col_bug_status = col_keyword + 3;
                ret.Add(new TestPlanKeyword(key, path, sheet, KeywordAtRow[key], col_keyword,
                    row_result, col_result, row_bug_status, col_bug_status, row_bug_list, col_bug_list));
            }
            return ret;
        }

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
    }
}
