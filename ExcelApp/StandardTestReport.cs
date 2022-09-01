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
            Excel.Application myReportExcel = ExcelAction.OpenOridnaryExcel(report_filename);
            if (myReportExcel == null)
            {
                Console.WriteLine("OpenOridnaryExcel failed in GenerateTestReportStructure()");
                return ret_testplan;
            }

            // Select and read Test Plan sheet
            Worksheet testplan_ws = ExcelAction.Find_Worksheet(myReportExcel, SheetName_TestPlan);
            if (testplan_ws == null)
            {
                Console.WriteLine("Find_Worksheet (TestPlan) failed in GenerateTestReportStructure()");
                return ret_testplan;
            }
            ret_testplan = TestPlan.LoadTestPlanSheet(testplan_ws);
            return ret_testplan;
        }

        public static void GenerateTestReportStructure(String report_filename)
        {
            List<TestPlan> read_testplan = new List<TestPlan>();

            read_testplan = ReadTestPlanFromStandardTestReport(report_filename);
            if (read_testplan == null) { return; }

            // Create a list of folder to be created and files to be copied (from/to)
            // filtered by Do or Not
            List<String> folder = new List<String>(), from = new List<String>(), to = new List<String>();
            foreach (TestPlan tp in read_testplan)
            {
                String group = tp.Group, summary = tp.Summary, do_or_not = tp.DoOrNot, subpart = tp.Subpart;
                if (do_or_not == "V")
                {
                    if (!folder.Contains(group)) { folder.Add(group); }
                    from.Add(@"Database Backup\" + summary.Substring(0, summary.IndexOf('_')) + ".xlsx");
                    to.Add(group + @"\" + summary + "_" + subpart + ".xlsx");
                }
            }

            // Create folder
            String input_report_dir = Path.GetDirectoryName(report_filename);
            String output_report_dir = input_report_dir + @"\TestReport_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            Directory.CreateDirectory(output_report_dir);
            foreach (String folder_name in folder)
            {
                Directory.CreateDirectory(output_report_dir + @"\" + folder_name);
            }

            // Copy files
            for (int index = 0; index < from.Count; index++)
            {
                String src = input_report_dir + @"\" + from[index];
                String dest = output_report_dir + @"\" + to[index];
                File.Copy(src, dest);
            }
        }

    }
}
