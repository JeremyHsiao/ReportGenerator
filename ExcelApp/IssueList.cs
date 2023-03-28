using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ExcelReportApplication
{
    public class IssueCount
    {
        enum SeverityOrder
        {
            A = 0,
            B,
            C,
            D
        };

        public static int severity_count = 4;
        private int[] count = new int[severity_count];
        public int Severity_A   // property
        {
            get { return count[(int)SeverityOrder.A]; }   // get method
            set { count[(int)SeverityOrder.A] = value; }  // set method
        }
        public int Severity_B   // property
        {
            get { return count[(int)SeverityOrder.B]; }   // get method
            set { count[(int)SeverityOrder.B] = value; }  // set method
        }
        public int Severity_C   // property
        {
            get { return count[(int)SeverityOrder.C]; }   // get method
            set { count[(int)SeverityOrder.C] = value; }  // set method
        }
        public int Severity_D   // property
        {
            get { return count[(int)SeverityOrder.D]; }   // get method
            set { count[(int)SeverityOrder.D] = value; }  // set method
        }

        public IssueCount() { count[0] = count[1] = count[2] = count[3] = 0;}
    }

    public class Issue
    {
        private String key;
        private String summary;
        private String severity;
        private String comment;
        private String status;
        private String reporter;
        private String assignee;
        private String duedate;
        private String testcaseid;
        private String bugtype;
        private String swversion;
        private String hwversion;
        private String linkedissue;
        private String additionalinfo;

        // out-of-band data
        private List<String> keyword_list;

        public String Key   // property
        {
            get { return key; }   // get method
            set { key = value; }  // set method
        }

        public String Summary   // property
        {
            get { return summary; }   // get method
            set { summary = value; }  // set method
        }

        public String Severity   // property
        {
            get { return severity; }   // get method
            set { severity = value; }  // set method
        }

        public String Comment   // property
        {
            get { return comment; }   // get method
            set { comment = value; }  // set method
        }

        public String Status   // property
        {
            get { return status; }   // get method
            set { status = value; }  // set method
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

        public String TestCaseID   // property
        {
            get { return testcaseid; }   // get method
            set { testcaseid = value; }  // set method
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

        public String LinkedIssue   // property
        {
            get { return linkedissue; }   // get method
            set { linkedissue = value; }  // set method
        }

        public String AdditionalInfo   // property
        {
            get { return additionalinfo; }   // get method
            set { additionalinfo = value; }  // set method
        }

        public List<String> KeywordList   // property
        {
            get { return keyword_list; }   // get method
            set { keyword_list = value; }  // set method
        }

        public const string col_Key = "Key";
        public const string col_Summary = "Summary";
        public const string col_Severity = "Severity";
        public const string col_RD_Comment = "Steps To Reproduce"; // used as comment currently 
        public const string col_Status = "Status"; 
        public const string col_Reporter = "Reporter"; 
        public const string col_Assignee = "Assignee";
        public const string col_DueDate = "Due Date";
        public const string col_TestCaseID = "Test Case ID";
        public const string col_BugType = "Bug Type";
        public const string col_SWVersion = "SW version";
        public const string col_HWVersion = "HW version";
        public const string col_LinkedIssue = "Linked Issues";
        public const string col_AdditionalInfo = "Additional Information"; 

        private void InitIssue() { keyword_list = new List<String>(); }

        public Issue()
        {
            InitIssue();
        }

        public Issue(String key, String summary, String severity, String comment
            , String status, String reporter, String assignee, String due, String tcid)
        {
            this.key = key; this.summary = summary; this.severity = severity; this.comment = comment;
            this.status = status; this.reporter = reporter; this.assignee = assignee; this.duedate = due; this.testcaseid = tcid;
            InitIssue();
        }

        public Issue(List<String> members)
        {
            this.key = members[(int)IssueListMemberIndex.KEY];
            this.summary = members[(int)IssueListMemberIndex.SUMMARY];
            this.severity = members[(int)IssueListMemberIndex.SEVERITY];
            this.comment = members[(int)IssueListMemberIndex.COMMENT];
            this.status = members[(int)IssueListMemberIndex.STATUS];
            this.reporter = members[(int)IssueListMemberIndex.REPORTER];
            this.assignee = members[(int)IssueListMemberIndex.ASSIGNEE];
            this.duedate = members[(int)IssueListMemberIndex.DUEDATE];
            this.testcaseid = members[(int)IssueListMemberIndex.TESTCASEID];
            this.bugtype = members[(int)IssueListMemberIndex.BUGTYPE];
            this.swversion = members[(int)IssueListMemberIndex.SWVERSION];
            this.hwversion = members[(int)IssueListMemberIndex.HWVERSION];
            this.linkedissue = members[(int)IssueListMemberIndex.LINKEDISSUE];
            this.additionalinfo = members[(int)IssueListMemberIndex.ADDITIONALINFO];
            InitIssue();
        }

        public enum IssueListMemberIndex
        {
            KEY = 0,
            SUMMARY,
            SEVERITY,
            COMMENT,
            STATUS,
            REPORTER,
            ASSIGNEE,
            DUEDATE,
            TESTCASEID,
            BUGTYPE,
            SWVERSION,
            HWVERSION,
            LINKEDISSUE,
            ADDITIONALINFO,
        }

        public static int IssueListMemberCount = Enum.GetNames(typeof(IssueListMemberIndex)).Length;

        // The sequence of this String[] must be aligned with enum IssueListMemberIndex (except no need to have string for MAX_NO)
        static String[] IssueListMemberColumnName = 
        { 
            col_Key,
            col_Summary,
            col_Severity,
            col_RD_Comment,
            col_Status,
            col_Reporter,
            col_Assignee,
            col_DueDate,
            col_TestCaseID,
            col_BugType,
            col_SWVersion,
            col_HWVersion,
            col_LinkedIssue,
            col_AdditionalInfo
        };

        static public String STR_CLOSE              = @"Close (0)";
        static public String STR_WAIVE              = @"Waive (0.1)";
        static public String STR_CONFIRM            = @"Confirm (1)";
        static public String STR_WFC                = @"WFC (2)";
        static public String STR_RD_ANALYSIS        = @"Analyzing and solving (3)";
        static public String STR_VENDOR_ANALYSIS    = @"Vendor analyzing (3.6)";
        static public String STR_MORE_INFO          = @"More Info. (3.9)";
        static public String STR_NEW                = @"New (4)";

        // constant strings for worksheet used in this application.
        static public string SheetName = "general_report";
        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
         // Key value
        static public string KeyPrefix = "BENSE";

        static public Color A_ISSUE_COLOR = Color.Red;
        static public Color B_ISSUE_COLOR = Color.DarkOrange;
        static public Color C_ISSUE_COLOR = Color.Black;

        static public List<Issue> GenerateIssueList(string buglist_filename)
        {
            List<Issue> ret_issue_list = new List<Issue>();

            ExcelAction.ExcelStatus status = ExcelAction.OpenIssueListExcel(buglist_filename);

            if (status == ExcelAction.ExcelStatus.OK)
            {
                Dictionary<string, int> col_name_list = ExcelAction.CreateIssueListColumnIndex();

                // Visit all rows and add content of IssueList
                int ExcelLastRow = ExcelAction.GetIssueListAllRange().Row;
                for (int excel_row_index = DataBeginRow; excel_row_index <= (ExcelLastRow - 1); excel_row_index++)  // Issue list until LastRow-1
                {
                    List<String> members = new List<String>();
                    for (int member_index = 0; member_index < IssueListMemberCount; member_index++)
                    {
                        String str;
                        // If data of xxx column exists in Excel, store it.
                        if (col_name_list.ContainsKey(IssueListMemberColumnName[member_index]))
                        {
                            str = ExcelAction.GetIssueListCellTrimmedString(excel_row_index, col_name_list[IssueListMemberColumnName[member_index]]);
                        }
                        // If not exist, fill an empty string to xxx
                        else
                        {
                            str = "";
                        }
                        members.Add(str);
                    }
                    // Add issue only if key contains KeyPrefix (very likely a valid key value)
                    if (members[(int)IssueListMemberIndex.KEY].Contains(KeyPrefix))
                    {
                        ret_issue_list.Add(new Issue(members));
                    }
                }
                ExcelAction.CloseIssueListExcel();
            }
            else
            {
                if (status == ExcelAction.ExcelStatus.ERR_OpenIssueListExcel_Find_Worksheet)
                {
                    // Worksheet not found -- data corruption -- need to check excel
                    ExcelAction.CloseIssueListExcel();
                }
                else
                {
                    // other error -- to be checked 
                }
            }

            return ret_issue_list;
        }

        static public Color descrption_color_issue = Color.Red;
        static public Color descrption_color_comment = Color.Blue;
         
        static public Dictionary<string, List<StyleString>> GenerateFullIssueDescription(List<Issue> issuelist)
        {
            Dictionary<string, List<StyleString>> ret_list = new Dictionary<string, List<StyleString>>();

            foreach (Issue issue in issuelist)
            {
                List<StyleString> value_style_str = new List<StyleString>();
                String key = issue.Key, rd_comment_str = issue.comment;

                if (key != "")
                {
                    String str = key + issue.Summary + "(" + issue.Severity + ")";
                    StyleString style_str = new StyleString(str, descrption_color_issue);
                    value_style_str.Add(style_str);

                    // Keep portion of string before first "\n"; if no "\n", keep whole string otherwise.
                    String short_comment = "";
                    if (rd_comment_str.Contains("\n"))
                    {
                        short_comment = rd_comment_str.Substring(0, rd_comment_str.IndexOf("\n"));
                    }
                    else
                    {
                        short_comment = rd_comment_str;
                    }
                    if (short_comment != "")
                    {
                        str = " --> " + short_comment;
                        style_str = new StyleString(str, descrption_color_comment);
                        value_style_str.Add(style_str);
                    }

                    // Add whole string into return_list
                    ret_list.Add(key, value_style_str);
                }
            }
            return ret_list;
        }
         
        // create key/rich-text-issue-description pair.
        // 
        // Format: KEY+SUMMARY+(+SEVERITY+)
        //
        // For example: BENSE27105-99[OSD]Menu scenario-Color Gamut value incorrect Without Metadata when Sub screen(B)
        //
        static public Dictionary<string, List<StyleString>> GenerateIssueDescription(List<Issue> issuelist)
        {
            Dictionary<string, List<StyleString>> ret_list = new Dictionary<string, List<StyleString>>();

            foreach (Issue issue in issuelist)
            {
                List<StyleString> value_style_str = new List<StyleString>();
                String key = issue.Key, rd_comment_str = issue.comment;

                if (key != "")
                {
                    String str = key + issue.Summary + "(" + issue.Severity + ")";
                    StyleString style_str = new StyleString(str, descrption_color_issue);
                    value_style_str.Add(style_str);
                    /*
                    // Keep portion of string before first "\n"; if no "\n", keep whole string otherwise.
                    String short_comment = "";
                    if (rd_comment_str.Contains("\n"))
                    {
                        short_comment = rd_comment_str.Substring(0, rd_comment_str.IndexOf("\n"));
                    }
                    else
                    {
                        short_comment = rd_comment_str;
                    }
                    if (short_comment != "")
                    {
                        str = " --> " + short_comment;
                        style_str = new StyleString(str, descrption_color_comment);
                        value_style_str.Add(style_str);
                    }
                    */
                    // Add whole string into return_list
                    ret_list.Add(key, value_style_str);
                }
            }
            return ret_list;
        }

        static public Dictionary<string, List<StyleString>> GenerateIssueDescription_Severity_by_Colors(List<Issue> issuelist)
        {
            Dictionary<string, List<StyleString>> ret_list = new Dictionary<string, List<StyleString>>();

            foreach (Issue issue in issuelist)
            {
                List<StyleString> value_style_str = new List<StyleString>();
                String key = issue.Key, rd_comment_str = issue.comment;

                if (key != "")
                {
                    Color color_by_severity;
                    switch (issue.Severity[0])
                    {
                        case 'A':
                            color_by_severity = Issue.A_ISSUE_COLOR;
                            break;
                        case 'B':
                            color_by_severity = Issue.B_ISSUE_COLOR;
                            break;
                        case 'C':
                            color_by_severity = Issue.C_ISSUE_COLOR;
                            break;
                        default:
                            color_by_severity = Color.Black;
                            break;
                    }
                    String str = key + issue.Summary + "(" + issue.Severity + ")";
                    StyleString style_str = new StyleString(str, color_by_severity);
                    value_style_str.Add(style_str);
                    // Add whole string into return_list
                    ret_list.Add(key, value_style_str);
                }
            }
            return ret_list;
        }

        // 
        // Input: keyword to check
        // Output: true: if contains keyword; false: not contain keyword
        // Note: Using this function so that it is easier to change the criteria of "Containing keyword"
        //
        public bool ContainKeyword(String Keyword)
        {
            bool ret = false;
            if ((this.Summary.Contains(Keyword)) || (this.TestCaseID.Contains(Keyword)))
            {
                ret = true;
            }
            return ret;
        }
    }
}
