using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReportApplication
{
    class TestReport
    {
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

            List<String> folder = new List<String> ();
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
            if (Storage.DirectoryExists(output_report_dir)) { return false; } // shouln't exist

            // Generate a list of all possible excel report files under input_report_dir -- SET A
            
            // Generate a list of excel report to be copied according to test-case file -- SET B

            // Copy files which belong to intersection (SET A, SET B) and record it (it is named SET A&B)

            // Record files which is availble in SET B nut not in SET A -- SET (B-A&B)
            // should be copied but not (because files don't exist in source

            // Record files which is availble in SET A but not in SET B -- SET (A-A&B)
            // files are not used in this time

            // go through all test-case and copy report files.
            foreach (TestCase tc in ReportGenerator.global_testcase_list)
            {
                String path = tc.Group, filename = tc.Summary + ".xlsx";
                String src_dir= report_Src + @"\" + path,
                       dest_dir = output_report_dir + @"\" + path;
                String src_file = Storage.GetValidFullFilename(src_dir,filename),
                      dest_file = Storage.GetValidFullFilename(output_report_dir,filename);

               if (Storage.FileExists(src_file))
               {
                   if (!Storage.DirectoryExists(output_report_dir))
                   {
                       Storage.CreateDirectory(output_report_dir, auto_parent_dir: true);
                   }
                   Storage.Copy(src_file, dest_file);
               }
            }

            return true;
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
