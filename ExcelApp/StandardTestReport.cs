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

    public class CopyTestReport
    {
        public String source_path;
        public String source_group;
        public String source_filename;
        public String destination_path;
        public String destination_group;
        public String destination_filename;
    }
    public enum Header_Update
    {
        Title = 1 << 0,
        Judgement = 1 << 1,
        Start_Day = 1 << 2,
        End_Day = 1 << 3,

        By_Template = 1 << 31,
    }

    class TestReport
    {
        public static int PassCnt_at_row = 21, PassCnt_at_col = 5;
        public static int FailCnt_at_row = 21, FailCnt_at_col = 7;
        public static int TotalCnt_at_row = 21, TotalCnt_at_col = 9;
        public static int Title_at_row = 1, Title_at_col = ('A' - 'A' + 1);
        public static int SW_Version_at_row = 7, SW_Version_at_col = ('J' - 'A' + 1);
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

        public static bool CopyTestReportbyExcelList(List<CopyTestReport> report_list)
        {
            //public String source_path;
            //public String source_group;
            //public String source_filename;
            //public String destination_path;
            //public String destination_group;
            //public String destination_filename;

            // Report to be copied according to List<CopyTestReport> report_list -- SET B (src_report_list)
            List<String> src_report_list = new List<String>();
            List<String> dest_report_list = new List<String>();
            foreach (CopyTestReport copy_report in report_list)
            {
                String src_path = Storage.CominePath(copy_report.source_path, copy_report.source_group);
                String src_file = copy_report.source_filename + ".xlsx";
                String src_fullfilename = Storage.GetValidFullFilename(src_path, src_file);
                if (Storage.FileExists(src_fullfilename))
                {
                    src_report_list.Add(src_fullfilename);

                    String dest_path = Storage.CominePath(copy_report.source_path, copy_report.source_group);
                    String dest_file = copy_report.source_filename + ".xlsx";
                    String dest_fullfilename = Storage.GetValidFullFilename(dest_path, dest_file);
                    if (Storage.FileExists(dest_fullfilename))
                    {
                        dest_report_list.Add(dest_fullfilename);
                    }
                }
            }

            // copy report files.
            List<String> report_actually_copied_list_src = new List<String>();
            List<String> report_actually_copied_list_dest = new List<String>();
            Dictionary<String, String> report_copied = CopyTestReport(src_report_list, dest_report_list);
            report_actually_copied_list_src = report_copied.Keys.ToList();
            report_actually_copied_list_dest = report_copied.Values.ToList();

            //AutoCorrectReport(

            return true;
        }

        //public static bool CopyTestReportbyTestCase(String report_Src, String output_report_dir)
        //{
        //    // Full file name exist checked before executing task

        //    String input_report_dir = report_Src;

        //    // test_plan (sample) dir must exist
        //    if (!Storage.DirectoryExists(input_report_dir)) { return false; }  // should exist

        //    // output test plan root_dir must be inexist so that no overwritten
        //    if (!Storage.DirectoryExists(output_report_dir)) { return false; } // should exist

        //    // Generate a list of all possible excel report files under input_report_dir -- SET A (all_report_list)
        //    List<String> src_dir_file_list = Storage.ListFilesUnderDirectory(report_Src);
        //    List<String> all_report_list = Storage.FilterFilename(src_dir_file_list);

        //    // Report to be copied according to test-case file -- SET B (src_report_list)
        //    List<String> src_report_list = new List<String>();
        //    // Files will be copied -- the intersection (SET A, SET B) (SET A&B report_can_be_copied_list_src)
        //    List<String> report_can_be_copied_list_src = new List<String>();
        //    List<String> report_can_be_copied_list_dest = new List<String>();
        //    // Files should be copied (listed in B) but not in A (because files don't exist in source SET A -- (B-A&B) report_not_available_list
        //    List<String> report_not_available_list = new List<String>();
        //    // files in A are not used in B this time -- SET (A-A&B) report_not_used_list
        //    List<String> report_not_used_list = new List<String>();

        //    // Set B is generated
        //    // Because A is also available, we can use it to generate A&B (check if this B is also in A) and B-A&B (check if this B is NOT in A)
        //    foreach (TestCase tc in ReportGenerator.global_testcase_list)
        //    {
        //        // go through all test-case and copy report files.
        //        String path = tc.Group, filename = tc.Summary + ".xlsx";
        //        String src_dir = Storage.CominePath(report_Src, path);
        //        String src_file = Storage.GetValidFullFilename(src_dir, filename);
        //        src_report_list.Add(src_file);          // item in SET B
        //        if (all_report_list.IndexOf(src_file) >= 0)
        //        {
        //            // also in A
        //            report_can_be_copied_list_src.Add(src_file);  // for A&B
        //            String dest_dir = Storage.CominePath(output_report_dir, path);
        //            String dest_file = Storage.GetValidFullFilename(dest_dir, filename);
        //            report_can_be_copied_list_dest.Add(dest_file);
        //        }
        //        else
        //        {
        //            // but not in A    
        //            report_not_available_list.Add(src_file);        // for B-A&B
        //        }
        //    }
        //    // SET (A-A&B) report_not_used_list
        //    report_not_used_list = all_report_list.Except(report_can_be_copied_list_src).ToList();
        //    // // Iterate A and add whatever not in A&B
        //    // foreach (String str in all_report_list)
        //    // {
        //    //     if(report_can_be_copied_list_src.IndexOf(str)<0)       // not in A&B
        //    //     {
        //    //         report_not_used_list.Add(str);
        //    //     }
        //    //}

        //    // copy report files.
        //    List<String> report_actually_copied_list_src = new List<String>();
        //    List<String> report_actually_copied_list_dest = new List<String>();
        //    Dictionary<String, String> report_copied = CopyTestReport(report_can_be_copied_list_src, report_can_be_copied_list_dest);
        //    report_actually_copied_list_src = report_copied.Keys.ToList();
        //    report_actually_copied_list_dest = report_copied.Values.ToList();
        //    //for(int index = 0; index < report_can_be_copied_list_src.Count; index++)
        //    //{
        //    //    String src_file = report_can_be_copied_list_src[index];
        //    //    if (Storage.FileExists(src_file))
        //    //    {
        //    //        String dest_file = report_can_be_copied_list_dest[index];
        //    //        string dest_path = Storage.GetDirectoryName(dest_file);
        //    //        if (!Storage.DirectoryExists(dest_path))
        //    //        {
        //    //            Storage.CreateDirectory(dest_path, auto_parent_dir: true);
        //    //        }
        //    //        if (Storage.Copy(src_file, dest_file, overwrite: true))
        //    //        {
        //    //            report_actually_copied_list_src.Add(src_file);
        //    //            report_actually_copied_list_dest.Add(dest_file);
        //    //        }
        //    //    }
        //    //    else
        //    //    {
        //    //        // cannot copy
        //    //    }
        //    //}

        //    // Clear Judgement of copied test_report
        //    UpdateAllHeader(report_actually_copied_list_dest, Judgement: " ");

        //    return true;
        //}

        public static bool CopyTestReportbyTestCase(String report_Src, String output_report_dir)
        {
            // Full file name exist checked before executing task

            String input_report_dir = report_Src;

            // test_plan (sample) dir must exist
            if (!Storage.DirectoryExists(input_report_dir)) { return false; }  // should exist

            // output test plan root_dir must be inexist so that no overwritten
            if (!Storage.DirectoryExists(output_report_dir)) { return false; } // should exist

            // Generate a list of all possible excel report files under input_report_dir -- SET A (all_report_list)
            List<String> all_report_list = Storage.ListCandidateReportFilesUnderDirectory(report_Src);

            // Report to be copied according to test-case file -- SET B (src_report_list)
            List<String> src_report_list = new List<String>();
            // Files will be copied -- the intersection (SET A, SET B) (SET A&B report_can_be_copied_list_src)
            Dictionary<String, String> report_to_be_copied = new Dictionary<String, String>();
            // Files should be copied (listed in B) but not in A (because files don't exist in source SET A -- (B-A&B) report_not_available_list
            List<String> report_not_available_list = new List<String>();
            // files in A are not used in B this time -- SET (A-A&B) report_not_used_list
            List<String> report_not_used_list = new List<String>();

            // Generating SET B (src_report_list)
            // Because A is also available, we can use it to generate A&B (check if this B is also in A) and B-A&B (check if this B is NOT in A)
            foreach (TestCase tc in ReportGenerator.global_testcase_list)
            {
                // go through all test-case and copy report files.
                String path = tc.Group, filename = tc.Summary + ".xlsx";
                String src_dir = Storage.CominePath(report_Src, path);
                String src_file = Storage.GetValidFullFilename(src_dir, filename);
                src_report_list.Add(src_file);          // item in SET B
                if (all_report_list.IndexOf(src_file) >= 0)   // if this item (of SET B) also in SET A ==? A&B
                {
                    String dest_dir = Storage.CominePath(output_report_dir, path);
                    String dest_file = Storage.GetValidFullFilename(dest_dir, filename);
                    report_to_be_copied.Add(src_file, dest_file);
                }
                else
                {
                    // but not in A    
                    report_not_available_list.Add(src_file);        // for B-A&B (in B but not in A)
                }
            }
            // SET (A-A&B) report_not_used_list -- this report is available under source path but not used for copying
            report_not_used_list = all_report_list.Except(report_to_be_copied.Keys).ToList();

            // copy report files.
            List<String> report_actually_copied_list_src = new List<String>();
            List<String> report_actually_copied_list_dest = new List<String>();
            // use Auto Correct Function to copy and auto-correct.

            foreach (String src in report_to_be_copied.Keys)
            {
                String dest = report_to_be_copied[src];
                if (CopyReportClearJudgement_SingleFile(source_file: src, destination_file: dest))
                {
                    report_actually_copied_list_src.Add(src);
                    report_actually_copied_list_dest.Add(dest);
                }
            }

            return true;
        }


        static public Boolean UpdateReportHeader(Worksheet ws, String Title = null, String SW_Version = null, String Test_Start = null, String Test_End = null,
                                String Judgement = null, String Template = null)
        {
            Boolean b_ret = false;
            // to-be-finished.
            if (Template != null)
            {

            }
            else
            {
                if (Title != null)
                {
                    ExcelAction.SetCellValue(ws, Title_at_row, Title_at_col, Title);
                }
                if (SW_Version != null)
                {
                    ExcelAction.SetCellValue(ws, SW_Version_at_row, SW_Version_at_col, Judgement);
                }
                if (Test_Start != null)
                {
                    ExcelAction.SetCellValue(ws, Period_Start_at_row, Period_Start_at_col, Test_Start);
                }
                if (Test_End != null)
                {
                    ExcelAction.SetCellValue(ws, Period_End_at_row, Period_End_at_col, Test_End);
                }
                if (Judgement != null)
                {
                    ExcelAction.SetCellValue(ws, Judgement_at_row, Judgement_at_col, Judgement);
                }
            }
            b_ret = true;
            return b_ret;
        }

        // Source report
        // Destination report (also update sheetname)
        // new report title
        // update header
        // clear judgement

        public static Boolean UpdateAllHeader(List<String> report_list, String Title = null, String SW_Version = null, String Test_Start = null, String Test_End = null,
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

                test_plan_status = plan.OpenDetailExcel(ReadOnly: false);
                if (test_plan_status == TestPlan.ExcelStatus.OK)
                {
                    UpdateReportHeader(plan.TestPlanWorksheet, Title: Title, SW_Version: SW_Version, Test_Start: Test_Start,
                                            Test_End: Test_End, Judgement: Judgement, Template: Template);
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
            Dictionary<String, String> Ret_copied_files = new Dictionary<String, String>();
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

        static public bool CopyReportClearJudgement_SingleFile(String source_file, String destination_file = "")
        {
            Boolean file_has_been_updated = false;

            if (Storage.IsReportFilename(source_file) == false)
            {
                // Do nothing if file does not look like a report file.
                return file_has_been_updated;
            }

            source_file = Storage.GetFullPath(source_file);
            destination_file = (destination_file != "") ? Storage.GetFullPath(destination_file) : source_file;

            // Open Excel workbook
            Workbook wb = ExcelAction.OpenExcelWorkbook(filename: source_file, ReadOnly: false);
            if (wb == null)
            {
                ConsoleWarning("ERR: Open workbook in AutoCorrectReport_SingleFile(): " + source_file);
                return false;
            }

            String sheet_name = TestPlan.GetSheetNameAccordingToFilename(source_file);
            Worksheet ws = ExcelAction.Find_Worksheet(wb, sheet_name);

            //// If valid sheet_name does not exist, use first worksheet and rename it.
            //String sheet_name = TestPlan.GetSheetNameAccordingToFilename(source_file);
            //Worksheet ws;
            //if (ExcelAction.WorksheetExist(wb, sheet_name) == false)
            //{
            //    ws = wb.Sheets[1];
            //    ws.Name = sheet_name;
            //    file_has_been_updated = true;
            //}
            //else
            //{
            //    ws = ExcelAction.Find_Worksheet(wb, sheet_name);
            //}

            // Update Judgement 
            TestReport.UpdateReportHeader(ws, Judgement: " ");

            // if parent directory does not exist, create recursively all parents
            String destination_dir = Storage.GetDirectoryName(destination_file);
            Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
            // Save
            ExcelAction.SaveExcelWorkbook(wb, filename: destination_file);
            file_has_been_updated = true;
            // Close Excel workbook
            ExcelAction.CloseExcelWorkbook(wb);

            return file_has_been_updated;
        }

        static public bool AutoCorrectReport_SingleFile(String source_file, String destination_file = "", Boolean always_save = false)
        {
            Boolean file_has_been_updated = false;

            if (Storage.IsReportFilename(source_file) == false)
            {
                // Do nothing if file does not look like a report file.
                return file_has_been_updated;
            }

            source_file = Storage.GetFullPath(source_file);
            destination_file = (destination_file != "") ? Storage.GetFullPath(destination_file) : source_file;

            // Open Excel workbook
            Workbook wb = ExcelAction.OpenExcelWorkbook(filename: source_file, ReadOnly: false);
            if (wb == null)
            {
                ConsoleWarning("ERR: Open workbook in AutoCorrectReport_SingleFile(): " + source_file);
                return false;
            }

            // If valid sheet_name does not exist, use first worksheet and rename it.
            String sheet_name = TestPlan.GetSheetNameAccordingToFilename(source_file);
            Worksheet ws;
            if (ExcelAction.WorksheetExist(wb, sheet_name) == false)
            {
                ws = wb.Sheets[1];
                ws.Name = sheet_name;
                file_has_been_updated = true;
            }
            else
            {
                ws = ExcelAction.Find_Worksheet(wb, sheet_name);
            }

            // Update header 
            String new_title = TestPlan.GetReportTitleAccordingToFilename(source_file);
            String existing_title = ExcelAction.GetCellTrimmedString(ws, TestReport.Title_at_row, TestReport.Title_at_col);
            if (existing_title != new_title)
            {
                TestReport.UpdateReportHeader(ws, Title: new_title);
                file_has_been_updated = true;
            }

            if ((file_has_been_updated) || (always_save))
            {
                // Something has been updated or always save (ex: to copy file & update) ==> save to excel file
                String destination_dir = Storage.GetDirectoryName(destination_file);
                // if parent directory does not exist, create recursively all parents
                Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
                ExcelAction.SaveExcelWorkbook(wb, filename: destination_file);
            }
            else
            {
                // Doing nothing here.
            }
            // Close Excel workbook
            ExcelAction.CloseExcelWorkbook(wb);

            return file_has_been_updated;
        }

        static public bool AutoCorrectReport(String report_root, String Output_dir = "")
        {
            Boolean b_ret = false;

            // 0.1 List all files under report_root_dir.
            List<String> file_list = Storage.ListFilesUnderDirectory(report_root);
            // 0.2 filename check to exclude non-report files.
            List<String> report_filename = Storage.FilterFilename(file_list);

            foreach (String source_report in report_filename)
            {
                String dest_filename = KeywordReport.DecideDestinationFilename(report_root, Output_dir, source_report);
                b_ret |= AutoCorrectReport_SingleFile(source_file: source_report, destination_file: dest_filename, always_save: true);
            }

            return b_ret;
        }

        //static public bool AutoCorrectReport(String report_root, String Output_dir = "")
        //{
        //    Boolean b_ret = false;

        //    // 0.1 List all files under report_root_dir.
        //    List<String> file_list = Storage.ListFilesUnderDirectory(report_root);
        //    // 0.2 filename check to exclude non-report files.
        //    List<String> report_filename = Storage.FilterFilename(file_list);

        //    //Output_dir = Storage.GetFullPath(Output_dir);

        //    // 1.1 Init an empty plan
        //    List<TestPlan> do_plan = new List<TestPlan>();

        //    // 1.2 Create a temporary test plan to includes report_file
        //    do_plan = TestPlan.CreateTempPlanFromFileList(report_filename);
        //    Boolean output_to_different_path = ((Output_dir=="")||(report_root==Output_dir))?false:true;

        //    foreach (TestPlan plan in do_plan)
        //    {
        //        String path = Storage.GetDirectoryName(plan.ExcelFile);
        //        String filename = Storage.GetFileName(plan.ExcelFile);
        //        String full_filename = Storage.GetFullPath(plan.ExcelFile);
        //        String sheet_name = plan.ExcelSheet;
        //        Boolean file_has_been_updated = output_to_different_path;

        //        // Open Excel workbood
        //        Workbook wb = ExcelAction.OpenExcelWorkbook(filename: full_filename, ReadOnly: false);
        //        if (wb == null)
        //        {
        //            ConsoleWarning("ERR: Open workbook in AutoCorrectReport(): " + full_filename);
        //            continue;
        //        }

        //        // If valid sheet_name does not exist, use first worksheet and rename it.
        //        Worksheet ws;
        //        if (ExcelAction.WorksheetExist(wb, sheet_name) == false)
        //        {
        //            ws = wb.Sheets[1];
        //            ws.Name = sheet_name;
        //            file_has_been_updated = true;
        //        }
        //        else
        //        {
        //            ws = ExcelAction.Find_Worksheet(wb, sheet_name);
        //        }

        //        // Update header 
        //        String new_title = TestPlan.GetReportTitleAccordingToFilename(filename);
        //        String existing_title = ExcelAction.GetCellTrimmedString(ws, TestReport.Title_at_row, TestReport.Title_at_col);
        //        if (existing_title != new_title)
        //        {
        //            TestReport.UpdateReportHeader(ws,Title: new_title);
        //            file_has_been_updated = true;
        //        }

        //        if (file_has_been_updated)
        //        {
        //            String dest_filename;
        //            // Something has been updated, save to excel file

        //            dest_filename = DecideDestinationFilename(report_root, Output_dir, full_filename);
        //            String dest_filename_dir = Storage.GetDirectoryName(dest_filename);
        //            // if parent directory does not exist, create recursively all parents
        //            Storage.CreateDirectory(dest_filename_dir, auto_parent_dir: true);
        //            ExcelAction.SaveExcelWorkbook(wb, filename: dest_filename);
        //            b_ret = true;
        //        }
        //        ExcelAction.CloseExcelWorkbook(wb);
        //    }
        //    return b_ret;
        //}

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
