using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReportApplication
{
    public class TestCase
    {
        private String key;
        private String group;
        private String summary;
        private String status;
        private String links;
        private String severity;
        private String bugtype;
        private String swversion;
        private String hwversion;
        private String reporter;
        private String assignee;
        private String duedate;
        private String additionalinfo;
        private String testcaseid;          
        private String stepstoreproduce;
        private String created;
        // generated-data
        private List<StyleString> linked_issue_list;
        private List<StyleString> keyword_issue_list;

        public String Key   // property
        {
            get { return key; }   // get method
            set { key = value; }  // set method
        }

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

        public String Status   // property
        {
            get { return status; }   // get method
            set { status = value; }  // set method
        }

        public String Links   // property
        {
            get { return links; }   // get method
            set { links = value; }  // set method
        }

        public String Severity   // property
        {
            get { return severity; }   // get method
            set { severity = value; }  // set method
        }

        public String BugType  // property
        {
            get { return bugtype; }   // get method
            set { bugtype = value; }  // set method
        }

        public String SWVersion   // property
        {
            get { return swversion; }   // get method
            set { swversion = value; }  // set method
        }

        public String HWVersion   // property
        {
            get { return hwversion; }   // get method
            set { hwversion = value; }  // set method
        }
       
        public String Reporter   // property
        {
            get { return reporter; }   // get method
            set { reporter = value; }  // set method
        }

        public String Assignee   // property
        {
            get { return assignee; }   // get method
            set { assignee = value; }  // set method
        }

        public String DueDate   // property
        {
            get { return duedate; }   // get method
            set { duedate = value; }  // set method
        }

        public String AdditionalInfo   // property
        {
            get { return additionalinfo; }   // get method
            set { additionalinfo = value; }  // set method
        }

        public String TestCaseID   // property
        {
            get { return testcaseid; }   // get method
            set { testcaseid = value; }  // set method
        }

        public String StepsToReproduce   // property
        {
            get { return stepstoreproduce; }   // get method
            set { stepstoreproduce = value; }  // set method
        }

        public String Created   // property
        {
            get { return created; }   // get method
            set { created = value; }  // set method
        }

        public List<StyleString> LinkedIssueList   // property
        {
            get { return linked_issue_list; }   // get method
            set { linked_issue_list = value; }  // set method
        }

        public List<StyleString> KeywordIssueList   // property
        {
            get { return keyword_issue_list; }   // get method
            set { keyword_issue_list = value; }  // set method
        }

        public const string col_Key = "Key";
        public const string col_Group = "Test Group";
        public const string col_Summary = "Summary";
        public const string col_Status = "Status";
        public const string col_LinkedIssue = "Linked Issues";
        public const string col_Severity = "Severity";
        public const string col_BugType = "Bug Type";
        public const string col_SWVersion = "SW version";
        public const string col_HWVersion = "HW version";
        public const string col_Reporter = "Reporter";
        public const string col_Assignee = "Assignee";
        public const string col_DueDate = "Due Date";
        public const string col_AdditionalInfo = "Additional Information";
        public const string col_TestCaseID = "Test Case ID";
        public const string col_StepsToReproduce = "Steps To Reproduce";
        public const string col_Created = "Created";

        public TestCase()
        {
        }
/*
        public TestCase(String key, String group, String summary, String status, String links)
        {
            this.key = key; this.group = group; this.summary = summary; this.status = status; this.links = links;
        }
*/
        public TestCase(List<String> members)
        {
            this.key = members[(int)TestCaseMemberIndex.KEY];
            this.group = members[(int)TestCaseMemberIndex.GROUP];
            this.summary = members[(int)TestCaseMemberIndex.SUMMARY];
            this.status = members[(int)TestCaseMemberIndex.STATUS];
            this.links = members[(int)TestCaseMemberIndex.LINKEDISSUE];
            this.severity = members[(int)TestCaseMemberIndex.SEVERITY];
            this.bugtype = members[(int)TestCaseMemberIndex.BUGTYPE];
            this.swversion = members[(int)TestCaseMemberIndex.SWVERSION];
            this.hwversion = members[(int)TestCaseMemberIndex.HWVERSION];
            this.reporter = members[(int)TestCaseMemberIndex.REPORTER];
            this.assignee = members[(int)TestCaseMemberIndex.ASSIGNEE];
            this.duedate = members[(int)TestCaseMemberIndex.DUEDATE];
            this.created = members[(int)TestCaseMemberIndex.CREATED];
            this.additionalinfo = members[(int)TestCaseMemberIndex.ADDITIONALINFO];
            this.testcaseid = members[(int)TestCaseMemberIndex.TESTCASEID];
            this.stepstoreproduce = members[(int)TestCaseMemberIndex.STEPSTOREPRODUCE];
        }

        public enum TestCaseMemberIndex
        {
            KEY = 0,
            GROUP,
            SUMMARY,
            SEVERITY,
            BUGTYPE,
            SWVERSION,
            HWVERSION,
            STATUS,
            REPORTER,
            ASSIGNEE,
            DUEDATE,
            CREATED,
            ADDITIONALINFO,
            TESTCASEID,
            LINKEDISSUE,
            STEPSTOREPRODUCE,
        }

        public static int TestCaseMemberCount = Enum.GetNames(typeof(TestCaseMemberIndex)).Length;

       // The sequence of this String[] must be aligned with enum TestCaseMemberIndex (except no need to have string for MAX_NO)
        static String[] TestCaseMemberColumnName = 
        { 
            col_Key,
            col_Group,
            col_Summary,
            col_Severity,
            col_BugType,
            col_SWVersion,
            col_HWVersion,
            col_Status,
            col_Reporter,
            col_Assignee,
            col_DueDate,
            col_Created,
            col_AdditionalInfo,
            col_TestCaseID,
            col_LinkedIssue,
            col_StepsToReproduce
        };

        static public String STR_FINISHED = @"Finished";
        static public String STR_BLOCKED = @"Blocked";
        static public String STR_TESTING = @"Testing";
        static public String STR_NONE = @"None";

        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
        static public string SheetName = "general_report";
        static public string KeyPrefix = "TCBEN";

        static private String[] separators = { "," };
        static public List<String> Convert_LinksString_To_ListOfString(String links)
        {
            List<String> ret_list = new List<String> ();
            // protection
            if ((links == null)||(links =="")) return ret_list;   // return empty new object
            // Separate keys into string[]
            String[] issues = links.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            if (issues == null) return ret_list;
            // string[] to List<String> (trimmed) and return
            foreach(String str in issues)
            {
                ret_list.Add(str.Trim());
            }
            return ret_list;
        }

        static public String Convert_ListOfString_To_LinkString(List<String> list)
        {
            String ret = "";
            // protection
            if (list == null) return ret;
            if (list.Count == 0) return ret;
            foreach (String str in list)
            {
                ret += str + ", ";
            }
            ret.Trim(); // remove " " at beginning & end
            if (ret[ret.Length-1] == ',') { ret.Remove(ret.Length - 1); }// remove last "," 
            return ret;
        }

        static public List<TestCase> GenerateTestCaseList(string tclist_filename)
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

            ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename);

            if (status == ExcelAction.ExcelStatus.OK)
            {
                Dictionary<string, int> tc_col_name_list = ExcelAction.CreateTestCaseColumnIndex();

                // Visit all rows and add content of TestCase
                int ExcelLastRow = ExcelAction.Get_Range_RowNumber(ExcelAction.GetTestCaseAllRange());
                for (int excel_row_index = DataBeginRow; excel_row_index <= ExcelLastRow; excel_row_index++)
                {
                    List<String> members = new List<String>();
                    for (int member_index = 0; member_index < TestCaseMemberCount; member_index++)
                    {
                        String str;
                        // If data of xxx column exists in Excel, store it.
                        if (tc_col_name_list.ContainsKey(TestCaseMemberColumnName[member_index]))
                        {
                            str = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, tc_col_name_list[TestCaseMemberColumnName[member_index]]);
                        }
                        // If not exist, fill an empty string to xxx
                        else
                        {
                            str = "";
                        }
                        members.Add(str);
                    }
                    // Add issue only if key contains KeyPrefix (very likely a valid key value)
                    if (members[(int)TestCaseMemberIndex.KEY].Contains(KeyPrefix))
                    {
                        ret_tc_list.Add(new TestCase(members));
                    }
                }
                ExcelAction.CloseTestCaseExcel();
            }
            else
            {
                if (status == ExcelAction.ExcelStatus.ERR_OpenTestCaseExcel_Find_Worksheet)
                {
                    // Worksheet not found -- data corruption -- need to check excel
                    ExcelAction.CloseTestCaseExcel();
                }
                else
                {
                    // other error -- to be checked 
                }
            }

            return ret_tc_list;
        }
    }
}
