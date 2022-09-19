using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelReportApplication
{
    static class UnderDevelopment
    {
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

        private void TestPlanKeywordInit() { issue_list = new List<String>(); tc_list = new List<String>(); }

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
    }

    public static class KeywordReport
    {
        static public List<TestPlanKeyword> ListAllKeyword(List<TestPlan> DoPlan)
        {
            List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
            foreach (TestPlan plan in DoPlan)
            {
                plan.OpenDetailExcel();
                List<TestPlanKeyword> plan_keyword = plan.ListKeyword();
                plan.CloseIssueListExcel();
                if (plan_keyword!=null)
                {
                    ret.AddRange(plan_keyword);
                }
            }
            return ret;
        }
    }

}
