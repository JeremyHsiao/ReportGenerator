using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

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
            return ret_testplan;
        }

        public static void GenerateTestReportStructure(List<TestPlan> read_testplan, String report_src_dir, String report_dest_dir)
        {
            // if testplan is not read yet, return
            if (read_testplan == null) { return; }
            // if src not exist, return
            if (!FileFunction.DirectoryExists(report_src_dir)) { return; }
            // dest_dir must be inexist
            if (FileFunction.DirectoryExists(report_dest_dir)) { return; }

            // Create a list of folder to be created and files to be copied (from/to)
            // filtered by Do or Not
            List<String> folder = new List<String>(), from = new List<String>(), to = new List<String>();
            foreach (TestPlan tp in read_testplan)
            {
                String group = tp.Group, summary = tp.Summary, do_or_not = tp.DoOrNot, subpart = tp.Subpart;
                if (do_or_not == "V")
                {
                    if (!folder.Contains(group)) { folder.Add(group); }
                    from.Add(summary.Substring(0, summary.IndexOf('_')) + ".xlsx");
                    to.Add(group + @"\" + summary + "_" + subpart + ".xlsx");
                }
            }

            Directory.CreateDirectory(report_dest_dir);
            foreach (String folder_name in folder)
            {
                Directory.CreateDirectory(report_dest_dir + @"\" + folder_name);
            }

            // Copy files
            for (int index = 0; index < from.Count; index++)
            {
                String src = report_src_dir + @"\" + from[index];
                String dest = report_dest_dir + @"\" + to[index];
                File.Copy(src, dest);
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

            String file_dir = FileFunction.GetDirectoryName(filename);
            String report_root_dir = FileFunction.GetDirectoryName(file_dir);
            String input_report_dir = report_root_dir + src_dir;
            String output_report_dir = FileFunction.GenerateDirectoryNameWithDateTime(report_root_dir);

            // test_plan (sample) dir must exist
            if (!FileFunction.DirectoryExists(input_report_dir)) { return false; }  // should exist

            // output test plan root_dir must be inexist so that no overwritten
            if (FileFunction.DirectoryExists(output_report_dir)) { return false; } // shouln't exist

            // read test-plan sheet NG and return if NG
            List<TestPlan> testplan = new List<TestPlan>();
            testplan = TestReport.ReadTestPlanFromStandardTestReport(filename);
            if (testplan == null) { return false; }

            // all input parameters has been checked successfully, so generate
            TestReport.GenerateTestReportStructure(testplan, input_report_dir, output_report_dir);
            return true;
        }

    }
}
