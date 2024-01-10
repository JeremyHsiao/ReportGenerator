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
        private String purpose;
        private String criteria;
        private String category;
        private String planver;

        // generated-data
        //private List<StyleString> linked_issue_list;
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

        public String Purpose   // property
        {
            get { return purpose; }   // get method
            set { purpose = value; }  // set method
        }

        public String Criteria   // property
        {
            get { return criteria; }   // get method
            set { criteria = value; }  // set method
        }

        public String Category   // property
        {
            get { return category; }   // get method
            set { category = value; }  // set method
        }

        public String PlanVer   // property
        {
            get { return planver; }   // get method
            set { planver = value; }  // set method
        }

        //public List<StyleString> LinkedIssueDescription   // property
        //{
        //    get { return linked_issue_list; }   // get method
        //    set { linked_issue_list = value; }  // set method
        //}

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
        public const string col_Purpose = "Test Case Purpose";
        public const string col_Criteria = "Test Case Criteria";
        public const string col_Category = "Test Case Category";
        public const string col_PlanVer = "Test Plan Ver.";

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
            this.purpose = members[(int)TestCaseMemberIndex.PURPOSE];
            this.criteria = members[(int)TestCaseMemberIndex.CRITERIA];
            this.category = members[(int)TestCaseMemberIndex.CATEGORY];
            this.planver = members[(int)TestCaseMemberIndex.PLANVER];
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
            PURPOSE,
            CRITERIA,
            CATEGORY,
            PLANVER
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
            col_StepsToReproduce,
            col_Purpose,
            col_Criteria,
            col_Category,
            col_PlanVer,
        };

        static public String STR_FINISHED = @"Finished";
        static public String STR_BLOCKED = @"Blocked";
        static public String STR_TESTING = @"Testing";
        static public String STR_NONE = @"None";

        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
        static public string SheetName = "general_report";
        static public string TemplateSheetName = "TC_List";
        static public string KeyPrefix = "T";

        static public void LoadFromXML()
        {
            // config for test-case
            KeyPrefix = XMLConfig.ReadAppSetting_String("TC_Key_Prefix");
            SheetName = XMLConfig.ReadAppSetting_String("TCList_ExportedSheetName");
            TemplateSheetName = XMLConfig.ReadAppSetting_String("TC_SheetName");
            NameDefinitionRow = XMLConfig.ReadAppSetting_int("TC_Row_NameDefine");
            DataBeginRow = XMLConfig.ReadAppSetting_int("TC_Row_DataBegin");
        }

        static public Boolean CheckValidTC_By_KeyPrefix(String tc_key)
        {
            Boolean ret = false;
            if ((tc_key.Length > TestCase.KeyPrefix.Length) && (String.Compare(tc_key, 0, TestCase.KeyPrefix, 0, TestCase.KeyPrefix.Length) == 0))
            {
                ret = true;
            }

            return ret;
        }

        static public Boolean CheckValidTC_By_Key_Summary(String tc_key, String tc_summary)
        {
            Boolean ret = false;
            if ((CheckValidTC_By_KeyPrefix(tc_key)) && (String.IsNullOrWhiteSpace(tc_summary) == false))
            //if ((tc_key.Length > TestCase.KeyPrefix.Length) &&
            //    (String.Compare(tc_key, 0, TestCase.KeyPrefix, 0, TestCase.KeyPrefix.Length) == 0) &&
            //    (String.IsNullOrWhiteSpace(tc_summary) == false))
            {
                ret = true;
            }

            return ret;
        }

        static public List<TestCase> GenerateTestCaseList_data_processing()
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

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
                String tc_key = members[(int)TestCaseMemberIndex.KEY];
                String summary = members[(int)TestCaseMemberIndex.SUMMARY];
                // Add issue only if key contains KeyPrefix (very likely a valid key value)
                //if (tc_key.Length < KeyPrefix.Length) { continue; } // If not a TC key in this row, go to next row
                //if (String.Compare(tc_key, 0, KeyPrefix, 0, KeyPrefix.Length) != 0) { continue; }
                //if (String.IsNullOrWhiteSpace(summary) == true) { continue; } // 2nd protection to prevent not a TC row

                //if (members[(int)TestCaseMemberIndex.KEY].Contains(KeyPrefix))
                if (CheckValidTC_By_Key_Summary(tc_key, summary))
                {
                    ret_tc_list.Add(new TestCase(members));
                }
            }

            return ret_tc_list;
        }



        // This is the version to be revised -- separate excel open/close away from data processing
        static public List<TestCase> GenerateTestCaseList_v2(string tclist_filename)
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

            Boolean tc_open = OpenTestCaseExcel(tclist_filename);

            if (tc_open)
            {
                ret_tc_list = GenerateTestCaseList_data_processing();
                Boolean status = CloseTestCaseExcel();
                ReportGenerator.UpdateGlobalTestcaseList(ret_tc_list);
                ReportGenerator.SetTestcaseLUT_by_Key(TestCase.UpdateTCListLUT_by_Key(ret_tc_list));
                ReportGenerator.SetTestcaseLUT_by_Sheetname(TestCase.UpdateTCListLUT_by_Sheetname(ret_tc_list));
            }
            else 
            {
                // other error -- to be checked 
            }

            return ret_tc_list;
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
                    String tc_key = members[(int)TestCaseMemberIndex.KEY];
                    String summary = members[(int)TestCaseMemberIndex.SUMMARY];
                    // Add issue only if key contains KeyPrefix (very likely a valid key value)
                    //if (tc_key.Length < KeyPrefix.Length) { continue; } // If not a TC key in this row, go to next row
                    //if (String.Compare(tc_key, 0, KeyPrefix, 0, KeyPrefix.Length) != 0) { continue; }
                    //if (String.IsNullOrWhiteSpace(summary) == true) { continue; } // 2nd protection to prevent not a TC row

                    //if (members[(int)TestCaseMemberIndex.KEY].Contains(KeyPrefix))
                    if (CheckValidTC_By_Key_Summary(tc_key, summary))
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

            ReportGenerator.UpdateGlobalTestcaseList(ret_tc_list);
            ReportGenerator.SetTestcaseLUT_by_Key(TestCase.UpdateTCListLUT_by_Key(ret_tc_list));
            ReportGenerator.SetTestcaseLUT_by_Sheetname(TestCase.UpdateTCListLUT_by_Sheetname(ret_tc_list));
            return ret_tc_list;
        }

        static public Boolean OpenTestCaseExcel(String tclist_filename)
        {
            ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename);
            if (status == ExcelAction.ExcelStatus.OK)
            {
                return true;
            }
            else if (status == ExcelAction.ExcelStatus.ERR_OpenTestCaseExcel_Find_Worksheet)
            {
                // Worksheet not found -- data corruption -- need to check excel
                status = ExcelAction.CloseTestCaseExcel();
                if (status != ExcelAction.ExcelStatus.OK)
                {
                    // To be debugged
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        static public Boolean OpenTCTemplateExcel(String tclist_filename)
        {
            ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename, IsTemplate: true);
            if (status == ExcelAction.ExcelStatus.OK)
            {
                return true;
            }
            else if (status == ExcelAction.ExcelStatus.ERR_OpenTestCaseExcel_Find_Worksheet)
            {
                // Worksheet not found -- data corruption -- need to check excel
                status = ExcelAction.CloseTestCaseExcel(IsTemplate: true);
                if (status != ExcelAction.ExcelStatus.OK)
                {
                    // To be debugged
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        static public Boolean CloseTestCaseExcel()
        {
            ExcelAction.ExcelStatus status = ExcelAction.CloseTestCaseExcel();
            if (status == ExcelAction.ExcelStatus.OK)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static public Boolean CloseTCTemplateExcel()
        {
            ExcelAction.ExcelStatus status = ExcelAction.CloseTestCaseExcel(IsTemplate: true);
            if (status == ExcelAction.ExcelStatus.OK)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        // Exce file is opened/closed by the caller function
        static public List<TestCase> GenerateTestCaseList_processing_data_New()
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

            ExcelData_testcase = ExcelAction.InitTCExcelData();

            // prepare LUT for member_index (from 0 in sequence) to column_index  
            List<int> LUT_member_to_column_index = new List<int>();
            foreach (String name in TestCaseMemberColumnName)
            {
                int index = ExcelData_testcase.Column_Name.IndexOf(name);
                LUT_member_to_column_index.Add(index);
            }

            // Visit all rows and add content of TestCase
            for (int line_index = 0; line_index < ExcelData_testcase.LineCount(); line_index++)
            {
                List<String> members = new List<String>();
                foreach (int column_index in LUT_member_to_column_index)
                {
                    String str = ExcelData_testcase.GetCell(line_index, column_index);
                    members.Add(str);
                }

                String tc_key = members[(int)TestCaseMemberIndex.KEY];
                String summary = members[(int)TestCaseMemberIndex.SUMMARY];
                if (CheckValidTC_By_Key_Summary(tc_key, summary))
                {
                    ret_tc_list.Add(new TestCase(members));
                }
            }

            return ret_tc_list;
        }

        static public List<TestCase> GenerateTestCaseList_New_v2(string tclist_filename)
        {
            List<TestCase> ret_tc_list = new List<TestCase>();
            Boolean tc_open = OpenTestCaseExcel(tclist_filename);

            if (tc_open)
            {
                ret_tc_list = GenerateTestCaseList_processing_data_New();
                Boolean status = CloseTestCaseExcel();
                ReportGenerator.UpdateGlobalTestcaseList(ret_tc_list);
                ReportGenerator.SetTestcaseLUT_by_Key(TestCase.UpdateTCListLUT_by_Key(ret_tc_list));
                ReportGenerator.SetTestcaseLUT_by_Sheetname(TestCase.UpdateTCListLUT_by_Sheetname(ret_tc_list));
            }
            else
            {
                // other error -- to be checked 
            }

            return ret_tc_list;
        }

        // This is the version to be revised -- separate excel open/close away from data processing
        static public ExcelData ExcelData_testcase;
        static public List<TestCase> GenerateTestCaseList_New(string tclist_filename)
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

            ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename);

            if (status == ExcelAction.ExcelStatus.OK)
            {
                ExcelData_testcase = ExcelAction.InitTCExcelData();

                // prepare LUT for member_index (from 0 in sequence) to column_index  
                List<int> LUT_member_to_column_index = new List<int>();
                foreach (String name in TestCaseMemberColumnName)
                {
                    int index = ExcelData_testcase.Column_Name.IndexOf(name);
                    LUT_member_to_column_index.Add(index);
                }

                // Visit all rows and add content of TestCase
                for (int line_index = 0; line_index < ExcelData_testcase.LineCount(); line_index++)
                {
                    List<String> members = new List<String>();
                    foreach (int column_index in LUT_member_to_column_index)
                    {
                        String str = ExcelData_testcase.GetCell(line_index, column_index);
                        members.Add(str);
                    }

                    String tc_key = members[(int)TestCaseMemberIndex.KEY];
                    String summary = members[(int)TestCaseMemberIndex.SUMMARY];
                    if (CheckValidTC_By_Key_Summary(tc_key, summary))
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

            ReportGenerator.UpdateGlobalTestcaseList(ret_tc_list);
            ReportGenerator.SetTestcaseLUT_by_Key(TestCase.UpdateTCListLUT_by_Key(ret_tc_list));
            ReportGenerator.SetTestcaseLUT_by_Sheetname(TestCase.UpdateTCListLUT_by_Sheetname(ret_tc_list));
            return ret_tc_list;
        }

        static public Dictionary<string, TestCase> UpdateTCListLUT_by_Key(List<TestCase> TC_list)
        {
            Dictionary<string, TestCase> ret_lut = new Dictionary<string, TestCase>();
            foreach (TestCase tc in TC_list)
            {
                if (ret_lut.ContainsKey(tc.Key) == true)
                {
                    continue;       // key already exists. shouldn't be here
                }
                ret_lut.Add(tc.Key, tc);
            }
            return ret_lut;
        }

        static public Dictionary<string, TestCase> UpdateTCListLUT_by_Sheetname(List<TestCase> TC_list)
        {
            Dictionary<string, TestCase> ret_lut = new Dictionary<string, TestCase>();
            foreach (TestCase tc in TC_list)
            {
                String sheetname = TestPlan.GetSheetNameAccordingToSummary(tc.Summary);
                if (ret_lut.ContainsKey(sheetname))
                {
                    // shouldn't be here. Excel needs to be fixed
                }
                else
                {
                    ret_lut.Add(sheetname, tc);
                }
            }
            return ret_lut;
        }

        //static public List<TestCase> UpdateTCLinkedIssueList(List<TestCase> tc_to_be_updated, List<Issue> issue_list_source,
        //                                        Dictionary<string, List<StyleString>> bug_description_list)
        //{
        //    List<TestCase> updated_tc = tc_to_be_updated;
        //    foreach (TestCase tc in updated_tc) // looping
        //    {
        //        String links = tc.Links;
        //        //if (links != "")
        //        if (String.IsNullOrWhiteSpace(links) == false)
        //        {
        //            List<StyleString> str_list;
        //            str_list = StyleString.FilteredBugID_to_BugDescription(links, issue_list_source, bug_description_list);
        //            tc.LinkedIssueDescription = str_list;
        //        }
        //    }
        //    return updated_tc;
        //}


    }
}
