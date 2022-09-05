using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

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
        private String stepstoreproduce ;

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
            this.additionalinfo = members[(int)TestCaseMemberIndex.ADDITIONALINFO];
            this.testcaseid = members[(int)TestCaseMemberIndex.TESTCASEID];
            this.stepstoreproduce = members[(int)TestCaseMemberIndex.STEPSTOREPRODUCE];
        }

        public enum TestCaseMemberIndex
        {
            KEY = 0,
            GROUP,
            SUMMARY,
            STATUS,
            LINKEDISSUE,
            SEVERITY,
            BUGTYPE,
            SWVERSION,
            HWVERSION,
            REPORTER,
            ASSIGNEE,
            DUEDATE,
            ADDITIONALINFO,
            TESTCASEID,
            STEPSTOREPRODUCE,
            MAX_NO
        }

       // The sequence of this String[] must be aligned with enum TestCaseMemberIndex (except no need to have string for MAX_NO)
        static String[] TestCaseMemberColumnName = 
        { 
            col_Key,
            col_Group,
            col_Summary,
            col_Status,
            col_LinkedIssue,
            col_Severity,
            col_BugType,
            col_SWVersion,
            col_HWVersion,
            col_Reporter,
            col_Assignee,
            col_DueDate,
            col_AdditionalInfo,
            col_TestCaseID,
            col_StepsToReproduce
        };

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

        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
        static public string SheetName = "general_report";
        static public string KeyPrefix = "TCBEN";

        static public List<TestCase> GenerateTestCaseList(string tclist_filename)
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

            // Open excel (read-only & corrupt-load)
            Excel.Application myTCExcel = ExcelAction.OpenPreviousExcel(tclist_filename);
            //Excel.Application myTCExcel = OpenOridnaryExcel(tclist_filename);
            if (myTCExcel != null)
            {
                Worksheet ws_tclist = ExcelAction.Find_Worksheet(myTCExcel, SheetName);
                if (ws_tclist != null)
                {
                    Dictionary<string, int> col_name_list = ExcelAction.CreateTableColumnIndex(ws_tclist, NameDefinitionRow);

                    // Get the last (row,col) of excel
                    Range rngLast = ws_tclist.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                    // Visit all rows and add content of TestCase
                    for (int index = DataBeginRow; index <= rngLast.Row; index++)
                    {
                        List<String> members = new List<String>();
                        for (int member_index = 0; member_index < (int)TestCaseMemberIndex.MAX_NO; member_index++)
                        {
                            Object cell_value2 = ws_tclist.Cells[index, col_name_list[TestCaseMemberColumnName[member_index]]].Value2;
                            String str = (cell_value2 == null) ? "" : cell_value2.ToString();
                            members.Add(str);
                        }
                        // Add issue only if key contains KeyPrefix (very likely a valid key value)
                        if (members[(int)TestCaseMemberIndex.KEY].Contains(KeyPrefix))
                        {
                            ret_tc_list.Add(new TestCase(members));
                        }
                    }
                }
                ExcelAction.CloseExcelWithoutSaveChanges(myTCExcel);
                myTCExcel = null;
            }
            return ret_tc_list;
        }
    }
}
