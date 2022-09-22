using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

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

        static public int NameDefinitionRow_TestPlan = 2;
        static public int DataBeginRow_TestPlan = 3;

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

        private const int col_indentifier = 2;
        private const int col_keyword = 3;
        private const int row_test_detail_start = 27;
        private const String identifier_str = "Item";

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
            for (int row_index = row_test_detail_start; row_index <= row_print_area; row_index++)
            {
                String cell_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, col_indentifier);
                if (cell_text == "") continue;
                if ((cell_text.Length > identifier_str.Length) &&
                    (cell_text.ToLowerInvariant().Contains(identifier_str.ToLowerInvariant())))
                {
                    cell_text = ExcelAction.GetCellTrimmedString(ws_testplan, row_index, col_keyword);
                    if (cell_text == "") { ConsoleWarning("Empty Keyword", row_index); continue; }
                    if (KeywordAtRow.ContainsKey(cell_text)) { ConsoleWarning("Duplicated Keyword", row_index); continue; }
                    KeywordAtRow.Add(cell_text, row_index);
                }
            }

            List<TestPlanKeyword> ret =  new List<TestPlanKeyword> ();
            foreach (String key in KeywordAtRow.Keys)
            {
                ret.Add(new TestPlanKeyword(key,path,sheet,KeywordAtRow[key],col_keyword));
            }
            return ret;
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

    public class TestPlanKeyword
    {
        private String keyword;
        private String workbook;
        private String worksheet;
        private int at_row;
        private int at_column;
        private List<String> issue_list;
        private List<String> tc_list;
        private List<StyleString> issue_description_list;
        private List<StyleString> tc_description_list;

        private void TestPlanKeywordInit()
        {
            issue_list = new List<String>(); tc_list = new List<String>();
            issue_description_list = new List<StyleString>(); tc_description_list = new List<StyleString>();
        }

        public TestPlanKeyword() { TestPlanKeywordInit(); }
        public TestPlanKeyword(String Keyword, String Workbook = "", String Worksheet = "", int AtRow = 0, int AtColumn = 0)
        {
            TestPlanKeywordInit();
            keyword = Keyword;
            workbook = Workbook;
            worksheet = Worksheet;
            at_row = AtRow;
            at_column = AtColumn;
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
            List<String> ret_str = new List<String> ();

            if (keyword != "")
            {
            }

            IssueList = ret_str;
        }

        public void UpdateIssueDescriptionList(List<StyleString> description)
        {
            List<StyleString> ret_str = new List<StyleString> ();

            if (IssueList == null)
            {
                UpdateIssueList();
            }

            if (IssueList != null)
            {
                
            }

            IssueDescriptionList = ret_str;
        }
    }

}
