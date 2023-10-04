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

        // newly-added
        private String customer;
        private String sw_version;
        private String hw_version;
        private String testplan_version;
        private String priority;

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

        public String Customer   // property
        {
            get { return customer; }   // get method
            set { customer = value; }  // set method
        }
        public String SW_Version   // property
        {
            get { return sw_version; }   // get method
            set { sw_version = value; }  // set method
        }
        public String HW_Version   // property
        {
            get { return hw_version; }   // get method
            set { hw_version = value; }  // set method
        }
        public String TestPlan_Version   // property
        {
            get { return testplan_version; }   // get method
            set { testplan_version = value; }  // set method
        }
        public String Priority   // property
        {
            get { return priority; }   // get method
            set { priority = value; }  // set method
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

        public Worksheet TestPlanWorksheet   // property
        {
            get { return ws_testplan; }   // get method
            set { ws_testplan = value; }  // set method
        }

        public TestPlan()
        {
        }

        public TestPlan(List<String> members)
        {
            this.SetByMembers(members);
        }

        public void SetByMembers(List<String> members)
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
        public const string col_Customer = "Customer";
        public const string col_SW_Version = "SW Version";
        public const string col_HW_Version = "HW Version";
        public const string col_TestPlan_Version = "Test Plan Ver.";
        public const string col_Priority = "Priority";
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
            int row_end = ExcelAction.Get_Range_RowNumber(rngLast);
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
            EX_SaveDetailExcel,
            EX_CloseDetailExcel,
            EX_SaveChangesAndCloseDetailExcel,
            MAX_NO
        };

        public ExcelStatus OpenDetailExcel(Boolean ReadOnly = true)
        {
            try
            {
                Workbook wb;

                // Open excel (read-only & corrupt-load)
                wb = ExcelAction.OpenExcelWorkbook(path, ReadOnly: ReadOnly);

                if (wb == null)
                {
                    wb_testplan = null;
                    ws_testplan = null;
                    return ExcelStatus.ERR_OpenDetailExcel_OpenExcelWorkbook;
                }

                Worksheet ws = ExcelAction.Find_Worksheet(wb, sheet);
                if (ws == null)
                {
                    wb_testplan = wb;
                    ws_testplan = null;
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

        public ExcelStatus SaveDetailExcel(String AsFilename)
        {
            try
            {
                ExcelAction.SaveExcelWorkbook(wb_testplan, AsFilename);
                return ExcelStatus.OK;
            }
            catch
            {
                return ExcelStatus.EX_SaveDetailExcel;
            }
        }

        public ExcelStatus CloseDetailExcel(Boolean SaveChanges = false, String AsFilename = "")
        {
            try
            {
                if (wb_testplan == null)
                {
                    return ExcelStatus.ERR_CloseDetailExcel_wb_null;
                }
                ExcelAction.CloseExcelWorkbook(wb_testplan, SaveChanges: false, AsFilename: AsFilename);
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

        public static String GetSheetNameAccordingToFilename(String filename)
        {
            String full_filename = Storage.GetFullPath(filename);
            String short_filename = Storage.GetFileName(full_filename);
            String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            String sheet_name = sp_str[0];
            return sheet_name;
        }

        public static String GetSheetNameAccordingToSummary(String summary)
        {
            String[] sp_str = summary.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            String sheet_name = sp_str[0];
            return sheet_name;
        }

        public static String GetReportTitleAccordingToFilename(String filename)
        {
            String full_filename = Storage.GetFullPath(filename);
            String short_filename_no_extension = Storage.GetFileNameWithoutExtension(full_filename);
            return short_filename_no_extension;
        }

        public static String GetReportTitleWithoutNumberAccordingToFilename(String filename)
        {
            String full_filename = Storage.GetFullPath(filename);
            String short_filename_no_extension = Storage.GetFileNameWithoutExtension(full_filename);
            String[] sp_str = short_filename_no_extension.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            String ret_string = short_filename_no_extension.Substring(sp_str[0].Length + 1);

            return ret_string;
        }

        public static TestPlan CreateTempPlanFromFile(String filename)
        {
            TestPlan ret_plan = new TestPlan();

            // File existing check protection (it is better also checked and giving warning before entering this function)
            if (Storage.FileExists(filename) != false)
            {
                // DoOrNot must be "V" & ExcelFile/ExcelSheet must be correct
                String full_filename = Storage.GetFullPath(filename);
                String short_filename = Storage.GetFileName(full_filename);
                String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                String sheet_name = sp_str[0];
                String subpart = sp_str[1];
                List<String> tp_str = new List<String>();
                tp_str.AddRange(new String[] { "N/A", short_filename, "N/A", "V", "N/A", subpart });
                ret_plan.ExcelFile = full_filename;
                ret_plan.ExcelSheet = sheet_name;
            }
            else
            {
                // no warning here, simply skip this file.
            }

            return ret_plan;
        }

        public static List<TestPlan> CreateTempPlanFromFileList(List<String> filename)
        {
            List<TestPlan> ret_plan = new List<TestPlan>();
            foreach (String name in filename)
            {
                // File existing check protection (it is better also checked and giving warning before entering this function)
                if (Storage.FileExists(name) == false)
                    continue; // no warning here, simply skip this file.

                // DoOrNot must be "V" & ExcelFile/ExcelSheet must be correct
                String full_filename = Storage.GetFullPath(name);
                String short_filename = Storage.GetFileName(full_filename);
                String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                String sheet_name = sp_str[0];
                String subpart = sp_str[1];
                List<String> tp_str = new List<String>();
                tp_str.AddRange(new String[] { "N/A", short_filename, "N/A", "V", "N/A", subpart });
                TestPlan tp = new TestPlan(tp_str);
                tp.ExcelFile = full_filename;
                tp.ExcelSheet = sheet_name;
                ret_plan.Add(tp);
            }
            return ret_plan;
        }

        //public Boolean UpdateTestReportHeader(String SW_Version = null, String Test_Start = null, String Test_End = null,
        //                                String Judgement = null, String Template = null)
        //{
        //    return TestReport.UpdateReportHeader(this.TestPlanWorksheet,SW_Version, Test_Start, Test_End, Judgement, Template);
        //}
        // Sorting Function
        //static public int Compare_Sheetname_Ascending_v01(String x, String y)
        //{
        //    int final_compare = 0;

        //    String sheetname_x = TestPlan.GetSheetNameAccordingToFilename(x);
        //    String sheetname_y = TestPlan.GetSheetNameAccordingToFilename(y);

        //    int sheetname_x_len = sheetname_x.Length;
        //    int sheetname_x_value_pos = sheetname_x.IndexOf('.');
        //    String x_str = sheetname_x.Substring(sheetname_x_value_pos + 1, sheetname_x_len - (sheetname_x_value_pos + 1));
        //    double x_value;
        //    Boolean x_is_double = double.TryParse(x_str, out x_value);

        //    int sheetname_y_len = sheetname_y.Length;
        //    int sheetname_y_value_pos = sheetname_y.IndexOf('.');
        //    String y_str = sheetname_y.Substring(sheetname_y_value_pos + 1, sheetname_y_len - (sheetname_y_value_pos + 1));
        //    double y_value;
        //    Boolean y_is_double = double.TryParse(y_str, out y_value);

        //    if (x_is_double == false)
        //    {
        //        if (y_is_double == true)
        //        {
        //            final_compare = 1;
        //        }
        //    }
        //    else if (y_is_double == false)
        //    {
        //        final_compare = -1;
        //    }
        //    // both are double, can be compared in value
        //    else if (x_value < y_value)
        //    {
        //        final_compare = -1;
        //    }
        //    else if (x_value > y_value)
        //    {
        //        final_compare = 1;
        //    }

        //    return final_compare;
        //}

        static public int Compare_Sheetname_Ascending(String sheetname_x, String sheetname_y)
        {
            int final_compare = 0;

            // process when one of sheetname is null
            if (sheetname_x == null)
            {
                if (sheetname_y == null)
                {
                    final_compare = 0;
                    return final_compare;
                }
                else
                {
                    final_compare = -1;
                    return final_compare;
                }
            }
            else if (sheetname_y == null)
            {
                final_compare = -1;
                return final_compare;
            }

            String[] subs_x = sheetname_x.Split('.');
            String[] subs_y = sheetname_y.Split('.');

            int compare_index = 0;

            while (true)
            {
                int x_value = 0, y_value = 0;

                Boolean x_no_more_point = (compare_index < subs_x.Count()) ? false : true;
                Boolean y_no_more_point = (compare_index < subs_y.Count()) ? false : true;

                // Comparison 1: reaching end of sheetname?
                if (x_no_more_point)
                {
                    if (y_no_more_point)
                    {
                        final_compare = 0;
                        break;
                    }
                    else
                    {
                        final_compare = -1;
                        break;
                    }
                }
                else if (y_no_more_point)
                {
                    final_compare = 1;
                    break;
                }

                String x_str = subs_x[compare_index];
                String y_str = subs_y[compare_index];

                Boolean x_is_value = Int32.TryParse(x_str, out x_value);
                Boolean y_is_value = Int32.TryParse(y_str, out y_value);

                // Comparison 2: comparing text vs value? default: value < text
                if (x_is_value == false)
                {
                    if (y_is_value == false)
                    {
                        final_compare = String.Compare(x_str, y_str);
                        if (final_compare != 0)
                        {
                            // break to return final_compare
                            break;
                        }
                        else
                        {
                            // maybe there are more points to compare 
                            compare_index++;
                        }
                    }
                    else
                    {
                        final_compare = 1;
                        break;
                    }
                }
                else if (y_is_value == false)
                {
                    final_compare = -1;
                    break;
                }
                else
                {

                    if (x_value < y_value)
                    {
                        final_compare = -1;
                        break;
                    }
                    else if (x_value > y_value)
                    {
                        final_compare = 1;
                        break;
                    }
                    else
                    {
                        // maybe there are more points to compare 
                        compare_index++; // go to compare next level
                    }
                }
            }
            return final_compare;
        }

        static public int Compare_Sheetname_by_Filename_Ascending(String filename_x, String filename_y)
        {

            String sheetname_x = TestPlan.GetSheetNameAccordingToFilename(filename_x);
            String sheetname_y = TestPlan.GetSheetNameAccordingToFilename(filename_y);

            return Compare_Sheetname_Ascending(sheetname_x, sheetname_y);
        }

        static public int Compare_Sheetname_by_Filename_Descending(String filename_x, String filename_y)
        {
            int compare_result_asceding = Compare_Sheetname_by_Filename_Ascending(filename_x, filename_y);
            return -compare_result_asceding;
        }
    }
}
