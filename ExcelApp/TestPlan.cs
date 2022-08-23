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

        enum TestPlanMemberIndex
        {
            GROUP = 0,
            SUMMARY,
            ASSIGNEE,
            DO_OR_NOT,
            CATEGORY,
            SUBPART,
            MAX_NO
        }
        // The sequence of this String[] must be aligned with enum TestPlanMemberIndex
        static String[] TestPlanMemberColumnName = { "Test Group", "Summary", "Assginee", "Do or Not", "Test Case Category", "Subpart" };

        static public int NameDefinitionRow_TestPlan = 2;
        static public int DataBeginRow_TestPlan = 3;

        public static List<TestPlan> LoadTestPlanSheet(Worksheet testplan_ws)
        {
            List<TestPlan> ret_testplan = new List<TestPlan>();

            // Create index for each column name
            Dictionary<string, int> col_name_list = ExcelAction.CreateTableColumnIndex(testplan_ws, NameDefinitionRow_TestPlan);

            // Get the last (row,col) of excel
            Range rngLast = testplan_ws.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

            // Visit all rows and add content 
            for (int index = DataBeginRow_TestPlan; index <= rngLast.Row; index++)
            {
                List<String> members = new List<String>();
                for (int member_index = 0; member_index < (int)TestPlanMemberIndex.MAX_NO; member_index++)
                {
                    Object cell_value2 = testplan_ws.Cells[index, col_name_list[TestPlanMemberColumnName[member_index]]].Value2;
                    String str = (cell_value2 == null) ? "" : cell_value2.ToString();
                    if (str == "")
                    {
                        break; // cannot be empty value; skip to next row
                    }
                    members.Add(str);
                }
                if (members.Count == (int)TestPlanMemberIndex.MAX_NO)
                {
                    TestPlan tp = new TestPlan(members);
                    ret_testplan.Add(tp);
                }
            }
            return ret_testplan;
        }
    }


}
