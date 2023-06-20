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

        public static String GetReportTitleAccordingToFilename(String filename)
        {
            String full_filename = Storage.GetFullPath(filename);
            String short_filename_no_extension = Storage.GetFileNameWithoutExtension(full_filename);
            return short_filename_no_extension;
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

        static private void ConsoleWarning(String function, int row, int col)
        {
            Console.WriteLine("Warning: please check " + function + " at (" + row.ToString() + "," + col.ToString() + ")");
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
}
