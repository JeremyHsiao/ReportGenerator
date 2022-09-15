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
            if (!FileFunction.DirectoryExists(out_root)) { return; }

            List<String> folder = new List<String> ();
            foreach (TestPlan tp in do_plan)
            {
                String dst = out_root + @"\" + tp.ExcelFile;
                String dst_dir = FileFunction.GetDirectoryName(dst);
                if (!FileFunction.DirectoryExists(dst_dir)) 
                { 
                    FileFunction.CreateDirectory(dst_dir);
                }

                String src = in_root + @"\" + tp.BackupSource;
                FileFunction.Copy(src, dst);
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
            List<TestPlan> testplan = TestReport.ReadTestPlanFromStandardTestReport(filename);
            if (testplan == null) { return false; }

            // all input parameters has been checked successfully, so generate
            List<TestPlan> do_plan = TestPlan.ListDoPlan(testplan);
            FileFunction.CreateDirectory(output_report_dir); // create output root-dir
            TestReport.GenerateTestReportStructure(do_plan, input_report_dir, output_report_dir);
            return true;
        }


        //
        // This Demo Identify Keyword on the excel and insert a column to list all issues containing that keyword
        //
        static int col_indentifier = 2;
        static int col_keyword = 3;
        static public bool KeywordIssueGenerationTask(string report_filename)
        {
            //
            // 1. Open Excel and find the sheet
            //

            String full_filename = FileFunction.GetFullPath(report_filename);
            String short_filename = FileFunction.GetFileName(full_filename);
            String sheet_name = short_filename.Substring(0, short_filename.IndexOf("_"));

            // File exist check is done outside

            Workbook wb_keyword_issue = ExcelAction.OpenExcelWorkbook(full_filename, ReadOnly: false);
            if (wb_keyword_issue == null)
            {
                ConsoleWarning("OpenExcelWorkbook in KeywordIssueGenerationTask");
                return false;
            }

            Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_keyword_issue, sheet_name);
            if (result_worksheet == null)
            {
                ConsoleWarning("Find_Worksheet in KeywordIssueGenerationTask");
                return false;
            }

            //
            // 2. Find out Printable Area
            //
            // Assummed that Printable area always starting at $A$1 (also data processing area)
            // So excel data processing area ends at Printable area (row_count,col_count)
            Range rngPrintable = ExcelAction.GetWorksheetPrintableRange(result_worksheet);
            int row_print_area = rngPrintable.Rows.Count;
            int column_print_area = rngPrintable.Columns.Count;

            //
            // 3. Find out all keywords and create LUT (keyword,row_index)
            //    output:  LUT (keyword,row_index)
            //
            const int row_test_detail_start = 27;
            const String identifier_str = "Item";
            // Read report file for keyword & its row and store into keyword/row dictionary
            // Search keyword within printable area
            Dictionary<String, int> KeywordAtRow = new Dictionary<String, int>();
            for (int row_index = row_test_detail_start; row_index <= row_print_area; row_index++)
            {
                String cell_text = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, col_indentifier);
                if (cell_text == "") continue;
                if ((cell_text.Length > identifier_str.Length) &&
                    (cell_text.ToLowerInvariant().Contains(identifier_str.ToLowerInvariant())))
                {
                    cell_text = ExcelAction.GetCellTrimmedString(result_worksheet, row_index, col_keyword);
                    if (cell_text == "") { ConsoleWarning("Empty Keyword", row_index); continue; }
                    if (KeywordAtRow.ContainsKey(cell_text)) { ConsoleWarning("Duplicated Keyword", row_index); continue; }
                    KeywordAtRow.Add(cell_text, row_index);
                }
            }

            //
            // 4. Use keyword to find out all issues that contains keyword. 
            //    put issue_id into a string contains many id separated by a comma ','
            //    then store this issue_id into LUT (keyword,ids)
            //    output: LUT (keyword,id_list)
            //
            Dictionary<String, List<String>> KeywordIssueIDList = new Dictionary<String, List<String>>();
            foreach (String keyword in KeywordAtRow.Keys)
            {
                List<String> id_list = new List<String>();
                foreach (IssueList issue in ReportGenerator.global_issue_list)
                {
                    if (issue.Summary.Contains(keyword))
                    {
                        id_list.Add(issue.Key);
                    }
                }
                KeywordIssueIDList.Add(keyword, id_list);
            }

            //
            // 5. input:  LUT (keyword,id_list) + LUT (id,color_desription) (from GenerateIssueDescription())
            //    output: LUT (keyword,color_desription_list)
            //         
            //    using: id_list -> ExtendIssueDescription() -> color_description_list
            // This issue description list is needed for keyword issue list
            ReportGenerator.global_issue_description_list = IssueList.GenerateIssueDescription(ReportGenerator.global_issue_list);

            // Go throught each keyword and turn id_list into color_description
            Dictionary<String, List<StyleString>> KeyWordIssueDescription = new Dictionary<String, List<StyleString>>();
            foreach (String keyword in KeywordAtRow.Keys)
            {
                List<String> id_list = KeywordIssueIDList[keyword];
                List<StyleString> issue_description = StyleString.ExtendIssueDescription(id_list, ReportGenerator.global_issue_description_list);
                KeyWordIssueDescription.Add(keyword, issue_description);
            }

            //
            // 6. input:  LUT (keyword,color_description_list) + LUT (id,row_index)
            //    output: write color_description_list at Excel(row_index,new_inserted_col outside printable area
            //         
            // Insert extra column just outside printable area.
            int insert_col = column_print_area + 1;
            ExcelAction.Insert_Column(result_worksheet, insert_col);

            foreach (String keyword in KeywordAtRow.Keys)
            {
                List<StyleString> issue_description = KeyWordIssueDescription[keyword];
                StyleString.WriteStyleString(result_worksheet, KeywordAtRow[keyword], insert_col, issue_description);
            }

            // Save as another file with yyyyMMddHHmmss
            string dest_filename = FileFunction.GenerateFilenameWithDateTime(full_filename);
            ExcelAction.CloseExcelWorkbook(wb_keyword_issue, SaveChanges: true, AsFilename: dest_filename);
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
