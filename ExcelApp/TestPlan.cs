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
    }


}
