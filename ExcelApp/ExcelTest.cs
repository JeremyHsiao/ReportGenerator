using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelReportApplication
{
    static class ExcelTest
    {
        // Assumed folder structure:
        // exection file dir: ./ 
        // report_file dir:   report file selected by user
        // test report src    ./SampleData/TestFileFolder/Database Backup
        // test report dest   ./SampleData/TestFileFolder/

        const string src  = @"\SampleData\TestFileFolder\Database Backup";
        const string dest = @"\SampleData\TestFileFolder\";
        public static bool ExcelTestMainTask(String filename)
        {
            // Check conditions before executing task
            
            // test_plan (sample) dir must exist
            String input_report_dir = FileFunction.GetCurrentDirectory() + src;
            if (!FileFunction.DirectoryExists(input_report_dir)) { return false; }
            // @"Database Backup\"

            // output test plan root_dir must be inexist so that no overwritten
            String output_report_dir = FileFunction.GenerateDirectoryNameWithDateTime(FileFunction.GetCurrentDirectory() + dest);
            if (FileFunction.DirectoryExists(output_report_dir)) { return false; }

            // main report_file must exist so that test-plan can be read
            String full_filename = FileFunction.GetFullPath(filename);
            if (!FileFunction.FileExists(full_filename)) { return false; }

            // read test-plan sheet NG and return if NG
            List<TestPlan> testplan = new List<TestPlan>();
            testplan = TestReport.ReadTestPlanFromStandardTestReport(full_filename);
            if (testplan == null) { return false; }

            // all input parameters has been checked successfully, so generate
            TestReport.GenerateTestReportStructure(testplan, input_report_dir, output_report_dir);
            return true;
        }
    }
}
