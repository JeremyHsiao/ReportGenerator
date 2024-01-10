using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Configuration;
using System.Text.RegularExpressions;

namespace ExcelReportApplication
{
    public class IssueCount
    {
        enum SeverityOrder
        {
            A = 0,
            B,
            C,
            D,
            Waived_A,
            Waived_B,
            Waived_C,
            Waived_D,
            Closed_A,
            Closed_B,
            Closed_C,
            Closed_D,
            COUNT
        };

        public static int severity_count = (int)SeverityOrder.COUNT;
        //public static int severity_count = Enum.GetNames(typeof(SeverityOrder)).Length;
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
        public int Waived_A   // property
        {
            get { return count[(int)SeverityOrder.Waived_A]; }   // get method
            set { count[(int)SeverityOrder.Waived_A] = value; }  // set method
        }
        public int Waived_B   // property
        {
            get { return count[(int)SeverityOrder.Waived_B]; }   // get method
            set { count[(int)SeverityOrder.Waived_B] = value; }  // set method
        }
        public int Waived_C   // property
        {
            get { return count[(int)SeverityOrder.Waived_C]; }   // get method
            set { count[(int)SeverityOrder.Waived_C] = value; }  // set method
        }
        public int Waived_D   // property
        {
            get { return count[(int)SeverityOrder.Waived_D]; }   // get method
            set { count[(int)SeverityOrder.Waived_D] = value; }  // set method
        }
        public int Closed_A   // property
        {
            get { return count[(int)SeverityOrder.Closed_A]; }   // get method
            set { count[(int)SeverityOrder.Closed_A] = value; }  // set method
        }
        public int Closed_B   // property
        {
            get { return count[(int)SeverityOrder.Closed_B]; }   // get method
            set { count[(int)SeverityOrder.Closed_B] = value; }  // set method
        }
        public int Closed_C   // property
        {
            get { return count[(int)SeverityOrder.Closed_C]; }   // get method
            set { count[(int)SeverityOrder.Closed_C] = value; }  // set method
        }
        public int Closed_D   // property
        {
            get { return count[(int)SeverityOrder.Closed_D]; }   // get method
            set { count[(int)SeverityOrder.Closed_D] = value; }  // set method
        }

        public IssueCount()
        {
            for (int index = 0; index < (int)SeverityOrder.COUNT; index++)
            {
                count[index] = 0;       // clear all
            }
        }

        public int TotalCount()
        {
            int total_count = 0;
            for (int index = 0; index < (int)SeverityOrder.COUNT; index++) // count all
            {
                total_count += count[index];
            }
            return total_count;
        }

        public int TotalWaived()
        {
            int total_count = 0;
            for (int index = (int)SeverityOrder.Waived_A; index <= (int)SeverityOrder.Waived_D; index++) // count all waived
            {
                total_count += count[index];
            }
            return total_count;
        }

        public int NotClosedCount()
        {
            int total_count = 0;
            for (int index = 0; index <= (int)SeverityOrder.Waived_D; index++) // count ABCD/Waived ABCD
            {
                total_count += count[index];
            }
            return total_count;
        }

        public int ABC_non_Wavied_IssueCount()
        {
            int total_count = 0;
            for (int index = 0; index <= (int)SeverityOrder.C; index++) // count ABC 
            {
                total_count += count[index];
            }
            return total_count;
        }

        static public IssueCount IssueListStatistic(List<Issue> issue_lust)
        {
            IssueCount ret_ic = new IssueCount();
            foreach (Issue issue in issue_lust)
            {
                if (issue.Status == Issue.STR_CLOSE)
                {
                    switch (issue.Severity[0])
                    {
                        case 'A':
                            ret_ic.Closed_A++;
                            break;
                        case 'B':
                            ret_ic.Closed_B++;
                            break;
                        case 'C':
                            ret_ic.Closed_C++;
                            break;
                        case 'D':
                            ret_ic.Closed_D++;
                            break;
                    }
                }
                else if (issue.Status == Issue.STR_WAIVE)
                {
                    switch (issue.Severity[0])
                    {
                        case 'A':
                            ret_ic.Waived_A++;
                            break;
                        case 'B':
                            ret_ic.Waived_B++;
                            break;
                        case 'C':
                            ret_ic.Waived_C++;
                            break;
                        case 'D':
                            ret_ic.Waived_D++;
                            break;
                    }
                }
                else // if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
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
        private List<String> testcaseid_list;

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

        private void InitIssue() { keyword_list = new List<String>(); testcaseid_list = new List<String>(); }

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

        private void SetupIssue(List<String> members)
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
            if (String.IsNullOrWhiteSpace(this.testcaseid) == false)
            {
                this.testcaseid_list = Split_String_To_ListOfString(testcaseid);
            }
        }

        public Issue(List<String> members)
        {
            SetupIssue(members);
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

        static public String STR_CLOSE = @"Close (0)";
        static public String STR_WAIVE = @"Waive (0.1)";
        static public String STR_CONFIRM = @"Confirm (1)";
        static public String STR_WFC = @"WFC (2)";
        static public String STR_RD_ANALYSIS = @"Analyzing and solving (3)";
        static public String STR_VENDOR_ANALYSIS = @"Vendor analyzing (3.6)";
        static public String STR_MORE_INFO = @"More Info. (3.9)";
        static public String STR_NEW = @"New (4)";

        // constant strings for worksheet used in this application.
        static public string SheetName = "general_report";
        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
        // Key value
        static public string KeyPrefix = "BENSE";

        static public string KeywordIssue_report_Font = "Gill Sans MT";
        static public int KeywordIssue_report_FontSize = 12;
        static public Color KeywordIssue_report_FontColor = System.Drawing.Color.Black;
        static public FontStyle KeywordIssue_report_FontStyle = FontStyle.Regular;
        static public Color KeywordIssue_A_ISSUE_COLOR = Color.Red;
        static public Color KeywordIssue_B_ISSUE_COLOR = Color.Black;
        static public Color KeywordIssue_C_ISSUE_COLOR = Color.Black;
        static public Color KeywordIssue_D_ISSUE_COLOR = Color.Black;
        static public Color KeywordIssue_WAIVED_ISSUE_COLOR = Color.Black;
        static public Color KeywordIssue_CLOSED_ISSUE_COLOR = Color.Black;

        static public void LoadFromXML()
        {
            // config for issue list
            KeyPrefix = XMLConfig.ReadAppSetting_String("Issue_Key_Prefix");
            SheetName = XMLConfig.ReadAppSetting_String("BugList_ExportedSheetName");
            NameDefinitionRow = XMLConfig.ReadAppSetting_int("Issue_Row_NameDefine");
            DataBeginRow = XMLConfig.ReadAppSetting_int("Issue_Row_DataBegin");

            // config for keyword report
            KeywordIssue_report_Font = XMLConfig.ReadAppSetting_String("KeywordIssue_report_Font");
            KeywordIssue_report_FontSize = XMLConfig.ReadAppSetting_int("KeywordIssue_report_FontSize");
            KeywordIssue_report_FontColor = XMLConfig.ReadAppSetting_Color("KeywordIssue_report_FontColor");
            KeywordIssue_report_FontStyle = XMLConfig.ReadAppSetting_FontStyle("KeywordIssue_report_FontStyle");
            KeywordIssue_A_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("KeywordIssue_A_Issue_Color");
            KeywordIssue_B_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("KeywordIssue_B_Issue_Color");
            KeywordIssue_C_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("KeywordIssue_C_Issue_Color");
            KeywordIssue_D_ISSUE_COLOR = XMLConfig.ReadAppSetting_Color("KeywordIssue_D_Issue_Color");
            KeywordIssue_WAIVED_ISSUE_COLOR = Issue.KeywordIssue_report_FontColor;
            KeywordIssue_CLOSED_ISSUE_COLOR = Issue.KeywordIssue_report_FontColor;
        }

        static public Boolean OpenBugListExcel(String buglist_filename)
        {
            ExcelAction.ExcelStatus status = ExcelAction.OpenIssueListExcel(buglist_filename);
            if (status == ExcelAction.ExcelStatus.OK)
            {
                return true;
            }
            else if (status == ExcelAction.ExcelStatus.ERR_OpenIssueListExcel_Find_Worksheet)
            {
                status = ExcelAction.CloseIssueListExcel();
                return false;
            }
            else
            {
                return false;
            }
         }

        static public Boolean CloseBugListExcel()
        {
            ExcelAction.ExcelStatus status = ExcelAction.CloseIssueListExcel();
            if (status == ExcelAction.ExcelStatus.OK)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static public List<Issue> GenerateIssueList_v2(string buglist_filename)
        {
            List<Issue> ret_issue_list = new List<Issue>();
            Boolean status = OpenBugListExcel(buglist_filename);

            if (status)
            {
                ret_issue_list = GenerateIssueList_processing_data();
                status = CloseBugListExcel();
                if (status==false)
                {
                    // To be debugged
                }
            }
            else
            {
                // To be debugged
            }
            return ret_issue_list;
        }

        static public List<Issue> GenerateIssueList_processing_data()
        {
            List<Issue> ret_issue_list = new List<Issue>();

            Dictionary<string, int> col_name_list = ExcelAction.CreateIssueListColumnIndex();
            List<String> for_checking_repeated_key = new List<String>();

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
                String key_str = members[(int)IssueListMemberIndex.KEY];
                //if (members[(int)IssueListMemberIndex.KEY].Contains(KeyPrefix))
                if ((String.IsNullOrWhiteSpace(key_str) == false) && (key_str.Contains(KeyPrefix)) && (for_checking_repeated_key.Contains(key_str) == false))
                {
                    ret_issue_list.Add(new Issue(members));
                    for_checking_repeated_key.Add(key_str);
                }
            }

            return ret_issue_list;
        }

        // This is the version to be revised -- separate excel open/close away from data processing
        static public List<Issue> GenerateIssueList(string buglist_filename)
        {
            List<Issue> ret_issue_list = new List<Issue>();

            ExcelAction.ExcelStatus status = ExcelAction.OpenIssueListExcel(buglist_filename);

            if (status == ExcelAction.ExcelStatus.OK)
            {
                Dictionary<string, int> col_name_list = ExcelAction.CreateIssueListColumnIndex();
                List<String> for_checking_repeated_key = new List<String>();

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
                    String key_str = members[(int)IssueListMemberIndex.KEY];
                    //if (members[(int)IssueListMemberIndex.KEY].Contains(KeyPrefix))
                    if ((String.IsNullOrWhiteSpace(key_str) == false) && (key_str.Contains(KeyPrefix)) && (for_checking_repeated_key.Contains(key_str) == false))
                    {
                        ret_issue_list.Add(new Issue(members));
                        for_checking_repeated_key.Add(key_str);
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

        static public Dictionary<string, Issue> UpdateIssueListLUT(List<Issue> issue_list)
        {
            Dictionary<string, Issue> ret_lut = new Dictionary<string, Issue>();
            foreach (Issue issue in issue_list)
            {
                if (ret_lut.ContainsKey(issue.Key) == true)
                {
                    continue;           // key are repeated. shouldn't be here
                }
                ret_lut.Add(issue.Key, issue);
            }
            return ret_lut;
        }

        static public List<Issue> KeyStringToListOfIssue(String issues_key_string, List<Issue> issue_list_source)
        {
            List<Issue> ret_list = new List<Issue>();
            List<String> issue_id_list = Issue.Split_String_To_ListOfString(issues_key_string);
            foreach (Issue issue in issue_list_source)
            {
                if (issue_id_list.IndexOf(issue.Key) >= 0)
                {
                    // issue found & added
                    ret_list.Add(issue);
                }
            }
            return ret_list;
        }

        static public List<String> ListOfIssueToListOfIssueKey(List<Issue> issue_list_source)
        {
            List<String> key_list = new List<String>();
            foreach (Issue issue in issue_list_source)
            {
                key_list.Add(issue.Key);
            }
            return key_list;
        }

        static public List<Issue> FilterIssueByStatus(List<Issue> issues_to_be_filtered, List<String> filter_list)
        {
            List<Issue> ret_list = new List<Issue>();
            foreach (Issue issue in issues_to_be_filtered)
            {
                if (filter_list.IndexOf(issue.Status) < 0)
                {
                    // not filtered status
                    ret_list.Add(issue);
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
            bool b_ret = false;

            String allowed_delimiter = @"/,;";                                    // slash, comma, semi-colon are allowed as delimiter
            String regexKeywordString = @"(?:[" + allowed_delimiter + @"]|^)\s*" +
                                        Regex.Escape(@Keyword) +                 // \QKeyword\E in regex
                                        @"\s*(?:[" + allowed_delimiter + @"]|$)";
            RegexStringValidator identifier_keyword_Regex = new RegexStringValidator(regexKeywordString);
            try
            {
                // Attempt validation. if regex is false (no keyword found at all) then jumping to catch(); 
                identifier_keyword_Regex.Validate(this.TestCaseID);
                b_ret = true;
            }
            catch (ArgumentException e)
            {
                // Validation failed.
            }
            return b_ret;
        }

        static private String[] separators = { "," };

        static public List<String> Split_String_To_ListOfString(String links)
        {
            List<String> ret_list = new List<String>();
            // protection
            //if ((links == null) || (links == "")) return ret_list;   // return empty new object
            if (String.IsNullOrWhiteSpace(links)) return ret_list;   // return empty new object
            // Separate keys into string[]
            String[] issues = links.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            if (issues == null) return ret_list;
            // string[] to List<String> (trimmed) and return
            foreach (String str in issues)
            {
                ret_list.Add(str.Trim());
            }
            return ret_list;
        }

        static public String Combine_ListOfString_to_String(List<String> list)
        {
            String ret = "";
            // protection
            if (list == null) return ret;
            if (list.Count == 0) return ret;
            foreach (String str in list)
            {
                ret += str + separators[0] + " ";
            }
            ret.Trim(); // remove " " at beginning & end
            if (ret[ret.Length - 1] == ',') { ret.Remove(ret.Length - 1); }// remove last "," 
            return ret;
        }


        static public Boolean severity_descending = true;
        static public Boolean key_descending = true;
        static public int Compare_Severity_then_Key(Issue x, Issue y)
        {
            String[] separators_key_value = { "-" };
            int final_compare = 0;

            // compare key value first
            int severity_compare = String.Compare(x.Severity, y.Severity);
            if (severity_compare == 0)
            {
                // same severity, then check key
                String[] key_x = x.Key.Split(separators_key_value, StringSplitOptions.RemoveEmptyEntries);
                String[] key_y = y.Key.Split(separators_key_value, StringSplitOptions.RemoveEmptyEntries);
                int x_value = Convert.ToInt32(key_x[1]);
                int y_value = Convert.ToInt32(key_y[1]);
                if (x_value > y_value)
                {
                    // descending means larger is earlier --> reverse the sequence --> reverse the result
                    final_compare = (key_descending) ? (-1) : (1);
                }
                else if (x_value < y_value)
                {
                    final_compare = (key_descending) ? (1) : (-1);
                }
                else
                {
                    final_compare = 0;
                }
            }
            else
            {
                // not the same severity
                // "smaller" char means higher severity ==> descending is smaller to larger --> no need to reveerse the result
                final_compare = (severity_descending) ? (severity_compare) : (-severity_compare);
            }

            return final_compare;
        }


        static public List<Issue> SortingBySeverityAndKey(List<Issue> issue_list, Boolean severity_descending = true, Boolean key_descending = true)
        {

            List<Issue> ret_list = new List<Issue>();
            ret_list.AddRange(issue_list);
            ret_list.Sort(Compare_Severity_then_Key);
            return ret_list;
        }
    }

    public class IssueList
    {
        private List<Issue> issue_list;
        private Dictionary<String, Issue> issue_lut_by_key;                     // related by key in bug-list
        private Dictionary<String, Issue> issue_lut_by_testcaseid;              // related by testcase id in bug-list
        private Dictionary<String, Issue> issue_lut_by_linkedissue;             // related by linked issue in bug-list

        public List<Issue> List   // property
        {
            get { return issue_list; }   // get method
            set { issue_list = value; }  // set method
        }

        public List<String> GetKeys     // property
        {
            get { return issue_lut_by_key.Keys.ToList(); }   // get method
        }

        public List<Issue> GetValues     // property
        {
            get { return issue_lut_by_key.Values.ToList(); }   // get method
        }

        public IssueList()
        {
            this.issue_list = new List<Issue>();
            this.issue_lut_by_key = new Dictionary<String, Issue>();
            this.issue_lut_by_testcaseid = new Dictionary<String, Issue>();
            this.issue_lut_by_linkedissue = new Dictionary<String, Issue>();
        }

        public void Clear()
        {
            this.issue_list.Clear();
            this.issue_lut_by_key.Clear();
            this.issue_lut_by_testcaseid.Clear();
            this.issue_lut_by_linkedissue.Clear();
        }
    }

}
