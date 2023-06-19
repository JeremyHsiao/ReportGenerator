using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReportApplication
{
    class TestReportHeader
    {
        private String title;
        private String model_name;
        private String panel_module;
        private String tcon_board;
        //
        //
        private String sw_version;
        private String test_stage;
        private String test_period;
        private String judgement;
    }

    class TestReport
    {
        public static int PassCnt_at_row = 21, PassCnt_at_col = 5;
        public static int FailCnt_at_row = 21, FailCnt_at_col = 7;
        public static int TotalCnt_at_row = 21, TotalCnt_at_col = 9;
        public static int Title_at_row = 1, Title_at_col = ('A' - 'A' + 1);
        public static int SW_Version_at_row = 7, SW_Version_at_col = ('J'-'A' + 1 );
        public static int Period_Start_at_row = 8, Period_Start_at_col = ('L' - 'A' + 1);
        public static int Period_End_at_row = 8, Period_End_at_col = ('M' - 'A' + 1);
        public static int Judgement_at_row = 9, Judgement_at_col = ('D' - 'A' + 1);
        public static int Judgement_string_at_row = 9, Judgement_string_at_col = 2;

        static List<TestPlan> global_tp = new List<TestPlan>();
        static public string SheetName_TestPlan = "Test Plan";

        public static List<TestPlan> ReadTestPlanFromStandardTestReport(String report_filename)
        {
            List<TestPlan> ret_testplan = new List<TestPlan>();

            // open standard test report
            Workbook wb_testplan = ExcelAction.OpenExcelWorkbook(report_filename);
            if (wb_testplan == null)
            {
                Console.WriteLine("OpenExcelWorkbook failed in GenerateTestReportStructure()");
                return ret_testplan;
            }

            // Select and read Test Plan sheet
            Worksheet result_ws = ExcelAction.Find_Worksheet(wb_testplan, SheetName_TestPlan);
            if (result_ws == null)
            {
                Console.WriteLine("Find_Worksheet (TestPlan) failed in GenerateTestReportStructure()");
                return ret_testplan;
            }
            ret_testplan = TestPlan.LoadTestPlanSheet(result_ws);

            ExcelAction.CloseExcelWorkbook(wb_testplan);

            // prepare addtional data (not part of Test Plan but required for accessing excel)
            foreach (TestPlan tp in ret_testplan)
            {
                String group = tp.Group, summary = tp.Summary, do_or_not = tp.DoOrNot, subpart = tp.Subpart;
                String sheet = summary.Substring(0, summary.IndexOf('_'));      // ex: A.1.1_OSD ==> A.1.1
                String src = sheet + ".xlsx";                                   // ex: A.1.1 + .xlsx ==> A.1.1.xlsx
                tp.BackupSource = src;
                String dst = group + @"\" + summary + "_" + subpart + ".xlsx";  // A.1.1_OSD ==> group\A.1.1_OSD_All.xlsx
                tp.ExcelFile = dst;
                tp.ExcelSheet = sheet;
            }

            return ret_testplan;
        }

        public static void GenerateTestReportStructure(List<TestPlan> do_plan, String in_root, String out_root)
        {
            // if testplan is not read yet, return
            if (do_plan == null) { return; }
            if (!Storage.DirectoryExists(out_root)) { return; }

            List<String> folder = new List<String>();
            foreach (TestPlan tp in do_plan)
            {
                String dst = out_root + @"\" + tp.ExcelFile;
                String dst_dir = Storage.GetDirectoryName(dst);
                if (!Storage.DirectoryExists(dst_dir))
                {
                    Storage.CreateDirectory(dst_dir);
                }

                String src = in_root + @"\" + tp.BackupSource;
                Storage.Copy(src, dst);
            }
        }

        // Assumed folder structure:
        // report_file_path:  report_root/0.0_DQA Test Report/FILENAME.xlsx
        // file_dir:          report_root/0.0_DQA Test Report
        // report_root_dir:   report_root/
        // input_report_dir:  report_root_dir/Database Backup
        // output_report_dir: report_root_dir_DATETIME

        const string src_dir = @"\Database Backup";
        public static bool CreateStandardTestReportTask(String filename)
        {
            // Full file name exist checked before executing task

            String file_dir = Storage.GetDirectoryName(filename);
            String report_root_dir = Storage.GetDirectoryName(file_dir);
            String input_report_dir = report_root_dir + src_dir;
            String output_report_dir = Storage.GenerateDirectoryNameWithDateTime(report_root_dir);

            // test_plan (sample) dir must exist
            if (!Storage.DirectoryExists(input_report_dir)) { return false; }  // should exist

            // output test plan root_dir must be inexist so that no overwritten
            if (Storage.DirectoryExists(output_report_dir)) { return false; } // shouln't exist

            // read test-plan sheet NG and return if NG
            List<TestPlan> testplan = TestReport.ReadTestPlanFromStandardTestReport(filename);
            if (testplan == null) { return false; }

            // all input parameters has been checked successfully, so generate
            List<TestPlan> do_plan = TestPlan.ListDoPlan(testplan);
            Storage.CreateDirectory(output_report_dir); // create output root-dir
            TestReport.GenerateTestReportStructure(do_plan, input_report_dir, output_report_dir);
            return true;
        }

        public static bool CopyTestReportbyTestCase(String report_Src, String output_report_dir)
        {
            // Full file name exist checked before executing task

            String input_report_dir = report_Src;

            // test_plan (sample) dir must exist
            if (!Storage.DirectoryExists(input_report_dir)) { return false; }  // should exist

            // output test plan root_dir must be inexist so that no overwritten
            if (!Storage.DirectoryExists(output_report_dir)) { return false; } // should exist

            // Generate a list of all possible excel report files under input_report_dir -- SET A (all_report_list)
            List<String> src_dir_file_list = Storage.ListFilesUnderDirectory(report_Src);
            List<String> all_report_list = Storage.FilterFilename(src_dir_file_list);

            // Report to be copied according to test-case file -- SET B (src_report_list)
            List<String> src_report_list = new List<String>();
            // Files will be copied -- the intersection (SET A, SET B) (SET A&B report_can_be_copied_list_src)
            List<String> report_can_be_copied_list_src = new List<String>();
            List<String> report_can_be_copied_list_dest = new List<String>();
            // Files should be copied (listed in B) but not in A (because files don't exist in source SET A -- (B-A&B) report_not_available_list
            List<String> report_not_available_list = new List<String>();
            // files in A are not used in B this time -- SET (A-A&B) report_not_used_list
            List<String> report_not_used_list = new List<String>();

            // Set B is generated
            // Because A is also available, we can use it to generate A&B (check if this B is also in A) and B-A&B (check if this B is NOT in A)
            foreach (TestCase tc in ReportGenerator.global_testcase_list)
            {
                // go through all test-case and copy report files.
                String path = tc.Group, filename = tc.Summary + ".xlsx";
                String src_dir = Storage.CominePath(report_Src, path);
                String src_file = Storage.GetValidFullFilename(src_dir, filename);
                src_report_list.Add(src_file);          // item in SET B
                if (all_report_list.IndexOf(src_file) >= 0)
                {
                    // also in A
                    report_can_be_copied_list_src.Add(src_file);  // for A&B
                    String dest_dir = Storage.CominePath(output_report_dir, path);
                    String dest_file = Storage.GetValidFullFilename(dest_dir, filename);
                    report_can_be_copied_list_dest.Add(dest_file);
                }
                else
                {
                    // but not in A    
                    report_not_available_list.Add(src_file);        // for B-A&B
                }
            }
            // SET (A-A&B) report_not_used_list
            report_not_used_list = all_report_list.Except(report_can_be_copied_list_src).ToList();
            // // Iterate A and add whatever not in A&B
            // foreach (String str in all_report_list)
            // {
            //     if(report_can_be_copied_list_src.IndexOf(str)<0)       // not in A&B
            //     {
            //         report_not_used_list.Add(str);
            //     }
            //}

            // copy report files.
            List<String> report_actually_copied_list_src = new List<String>();
            List<String> report_actually_copied_list_dest = new List<String>();
            Dictionary<String,String> report_copied = CopyTestReport(report_can_be_copied_list_src, report_can_be_copied_list_dest);
            report_actually_copied_list_src = report_copied.Keys.ToList();
            report_actually_copied_list_dest = report_copied.Values.ToList();
            //for(int index = 0; index < report_can_be_copied_list_src.Count; index++)
            //{
            //    String src_file = report_can_be_copied_list_src[index];
            //    if (Storage.FileExists(src_file))
            //    {
            //        String dest_file = report_can_be_copied_list_dest[index];
            //        string dest_path = Storage.GetDirectoryName(dest_file);
            //        if (!Storage.DirectoryExists(dest_path))
            //        {
            //            Storage.CreateDirectory(dest_path, auto_parent_dir: true);
            //        }
            //        if (Storage.Copy(src_file, dest_file, overwrite: true))
            //        {
            //            report_actually_copied_list_src.Add(src_file);
            //            report_actually_copied_list_dest.Add(dest_file);
            //        }
            //    }
            //    else
            //    {
            //        // cannot copy
            //    }
            //}

            // Clear Judgement of copied test_report
            UpdateHeader(report_actually_copied_list_dest, Judgement:" ");

            return true;
        }

        // Source report
        // Destination report (also update sheetname)
        // new report title
        // update header
        // clear judgement

        public static Boolean UpdateHeader(List<String> report_list, String SW_Version = null, String Test_Start= null, String Test_End = null, 
                                        String Judgement = null, String Template = null)
        {
            // Create a temporary test plan to includes all files listed in List<String> report_filename
            List<TestPlan> do_plan = TestPlan.CreateTempPlanFromFileList(report_list);

            foreach (TestPlan plan in do_plan)
            {
                String path = Storage.GetDirectoryName(plan.ExcelFile);
                String filename = Storage.GetFileName(plan.ExcelFile);
                String sheet_name = plan.ExcelSheet;
                TestPlan.ExcelStatus test_plan_status;

                test_plan_status = plan.OpenDetailExcel(ReadOnly:false);
                if (test_plan_status == TestPlan.ExcelStatus.OK)
                {
                    // to-be-finished.
                    if (Template != null)
                    {

                    }
                    else
                    {
                        if (SW_Version != null)
                        {
                            ExcelAction.SetCellValue(plan.TestPlanWorksheet, SW_Version_at_row, SW_Version_at_col, Judgement);
                        }
                        if (Test_Start != null)
                        {
                            ExcelAction.SetCellValue(plan.TestPlanWorksheet, Period_Start_at_row, Period_Start_at_col, Test_Start);
                        }
                        if (Test_End != null)
                        {
                            ExcelAction.SetCellValue(plan.TestPlanWorksheet, Period_End_at_row, Period_End_at_col, Test_End);
                        }
                        if (Judgement != null)
                        {
                            ExcelAction.SetCellValue(plan.TestPlanWorksheet, Judgement_at_row, Judgement_at_col, Judgement);
                        }
                    }
                    plan.SaveDetailExcel(plan.ExcelFile);
                    plan.CloseDetailExcel();
                }
            }
            return true;
        }

        public static Dictionary<String, String> CopyTestReport(List<String> src_list, List<String> dest_list)
        {
            var dic = src_list.Zip(dest_list, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
            Dictionary<String, String> Ret_copied_files = CopyTestReport(dic);
            return Ret_copied_files;
        }

        public static Dictionary<String, String> CopyTestReport(Dictionary<String, String> src_dest_file_list)
        {
            Dictionary<String, String> Ret_copied_files = new Dictionary<String, String> ();
            foreach (String src_file in src_dest_file_list.Keys)
            {
                if (Storage.FileExists(src_file))
                {
                    String dest_file = src_dest_file_list[src_file];
                    string dest_path = Storage.GetDirectoryName(dest_file);
                    if (!Storage.DirectoryExists(dest_path))
                    {
                        Storage.CreateDirectory(dest_path, auto_parent_dir: true);
                    }
                    if (Storage.Copy(src_file, dest_file, overwrite: true))
                    {
                        Ret_copied_files.Add(src_file, dest_file);
                    }
                }
                else
                {
                    // file not exist
                }
            }

            return Ret_copied_files;
        }

        //static public ReadReport(String root_dir)
        //{
        //}

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
