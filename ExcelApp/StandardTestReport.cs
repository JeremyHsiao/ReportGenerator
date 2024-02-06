using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReportApplication
{
    class StandardTestReportHeader
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

    public enum StandardTestREportHeader_Update
    {
        Title = 1 << 0,
        Judgement = 1 << 1,
        Start_Day = 1 << 2,
        End_Day = 1 << 3,

        By_Template = 1 << 31,
    }

    class StandareTestReport
    {
        static List<TestPlan> global_tp = new List<TestPlan>();
        static public string SheetName_TestPlan = "Test Plan";
/*
        public static List<TestPlan> ReadTestPlanFromStandardTestReport(String report_filename)
        {
            List<TestPlan> ret_testplan = new List<TestPlan>();

            // open standard test report
            Workbook wb_testplan = ExcelAction.OpenExcelWorkbook(report_filename);
            if (wb_testplan == null)
            {
                LogMessage.WriteLine("OpenExcelWorkbook failed in GenerateTestReportStructure()");
                return ret_testplan;
            }

            // Select and read Test Plan sheet
            Worksheet result_ws = ExcelAction.Find_Worksheet(wb_testplan, SheetName_TestPlan);
            if (result_ws == null)
            {
                LogMessage.WriteLine("Find_Worksheet (TestPlan) failed in GenerateTestReportStructure()");
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
*/
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
                if (Storage.DirectoryExists(dst_dir) == false)
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

        //const string src_dir = @"\Database Backup";
        //public static bool CreateStandardTestReportTask(String filename)
        //{
        //    // Full file name exist checked before executing task

        //    String file_dir = Storage.GetDirectoryName(filename);
        //    String report_root_dir = Storage.GetDirectoryName(file_dir);
        //    String input_report_dir = report_root_dir + src_dir;
        //    String output_report_dir = Storage.GenerateDirectoryNameWithDateTime(report_root_dir);

        //    // test_plan (sample) dir must exist
        //    if (!Storage.DirectoryExists(input_report_dir)) { return false; }  // should exist

        //    // output test plan root_dir must be inexist so that no overwritten
        //    if (Storage.DirectoryExists(output_report_dir)) { return false; } // shouln't exist

        //    // read test-plan sheet NG and return if NG
        //    List<TestPlan> testplan = TestReport.ReadTestPlanFromStandardTestReport(filename);
        //    if (testplan == null) { return false; }

        //    // all input parameters has been checked successfully, so generate
        //    List<TestPlan> do_plan = TestPlan.ListDoPlan(testplan);
        //    Storage.CreateDirectory(output_report_dir); // create output root-dir
        //    TestReport.GenerateTestReportStructure(do_plan, input_report_dir, output_report_dir);
        //    return true;
        //}

        //public static bool CopyTestReportbyExcelList(List<CopyTestReport> report_list)
        //{
        //    //public String source_path;
        //    //public String source_folder;
        //    //public String source_group;
        //    //public String source_filename;
        //    //public String destination_path;
        //    //public String destination_folder;
        //    //public String destination_group;
        //    //public String destination_filename;

        //    Dictionary<String,String> copy_list = new Dictionary<String,String>();
        //    foreach (CopyTestReport copy_report in report_list)
        //    {
        //        String src_path = copy_report.Get_SRC_Directory();
        //        String src_fullfilename = copy_report.Get_SRC_FullFilePath();
        //        if (!Storage.FileExists(src_fullfilename))
        //            continue;

        //        String dest_path = copy_report.Get_DEST_Directory();
        //        String dest_fullfilename = copy_report.Get_DEST_FullFilePath();
        //        copy_list.Add(src_fullfilename,dest_fullfilename);
        //    }

        //    // copy report files.
        //    List<String> report_actually_copied_list_src = new List<String>();
        //    List<String> report_actually_copied_list_dest = new List<String>();
        //    // use Auto Correct Function to copy and auto-correct.

        //    foreach (String src in copy_list.Keys)
        //    {
        //        String dest = copy_list[src];
        //        if (AutoCorrectReport_SingleFile(source_file: src, destination_file: dest))
        //        {
        //            report_actually_copied_list_src.Add(src);
        //            report_actually_copied_list_dest.Add(dest);
        //        }
        //    }

        //    if (report_actually_copied_list_src.Count > 0)
        //        return true;
        //    else
        //        return false;
        //}

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

        //public static bool CopyTestReportbyTestCase(String report_Src, String output_report_dir)
        //{
        //    // Full file name exist checked before executing task

        //    String input_report_dir = report_Src;

        //    // test_plan (sample) dir must exist
        //    if (!Storage.DirectoryExists(input_report_dir)) { return false; }  // should exist

        //    // output test plan root_dir must be inexist so that no overwritten
        //    if (!Storage.DirectoryExists(output_report_dir)) { return false; } // should exist

        //    // Generate a list of all possible excel report files under input_report_dir -- SET A (all_report_list)
        //    List<String> all_report_list = Storage.ListCandidateReportFilesUnderDirectory(report_Src);

        //    // Report to be copied according to test-case file -- SET B (src_report_list)
        //    List<String> src_report_list = new List<String>();
        //    // Files will be copied -- the intersection (SET A, SET B) (SET A&B report_can_be_copied_list_src)
        //    Dictionary<String, String> report_to_be_copied = new Dictionary<String, String>();
        //    // Files should be copied (listed in B) but not in A (because files don't exist in source SET A -- (B-A&B) report_not_available_list
        //    List<String> report_not_available_list = new List<String>();
        //    // files in A are not used in B this time -- SET (A-A&B) report_not_used_list
        //    List<String> report_not_used_list = new List<String>();

        //    // Generating SET B (src_report_list)
        //    // Because A is also available, we can use it to generate A&B (check if this B is also in A) and B-A&B (check if this B is NOT in A)
        //    foreach (TestCase tc in ReportGenerator.ReadGlobalTestcaseList())
        //    {
        //        // go through all test-case and copy report files.
        //        String path = tc.Group, filename = tc.Summary + ".xlsx";
        //        String src_dir = Storage.CominePath(report_Src, path);
        //        String src_file = Storage.GetValidFullFilename(src_dir, filename);
        //        src_report_list.Add(src_file);          // item in SET B
        //        if (all_report_list.IndexOf(src_file) >= 0)   // if this item (of SET B) also in SET A ==? A&B
        //        {
        //            String dest_dir = Storage.CominePath(output_report_dir, path);
        //            String dest_file = Storage.GetValidFullFilename(dest_dir, filename);
        //            report_to_be_copied.Add(src_file, dest_file);
        //        }
        //        else
        //        {
        //            // but not in A    
        //            report_not_available_list.Add(src_file);        // for B-A&B (in B but not in A)
        //        }
        //    }
        //    // SET (A-A&B) report_not_used_list -- this report is available under source path but not used for copying
        //    report_not_used_list = all_report_list.Except(report_to_be_copied.Keys).ToList();

        //    // copy report files.
        //    List<String> report_actually_copied_list_src = new List<String>();
        //    List<String> report_actually_copied_list_dest = new List<String>();
        //    // use Auto Correct Function to copy and auto-correct.

        //    foreach (String src in report_to_be_copied.Keys)
        //    {
        //        String dest = report_to_be_copied[src];
        //        if (CopyReportClearJudgement_SingleFile(source_file: src, destination_file: dest))
        //        {
        //            report_actually_copied_list_src.Add(src);
        //            report_actually_copied_list_dest.Add(dest);
        //        }
        //    }

        //    if (report_actually_copied_list_src.Count > 0)
        //        return true;
        //    else
        //        return false;
        //}

        //public static Dictionary<String, String> CopyTestReport(List<String> src_list, List<String> dest_list)
        //{
        //    var dic = src_list.Zip(dest_list, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
        //    Dictionary<String, String> Ret_copied_files = CopyTestReport(dic);
        //    return Ret_copied_files;
        //}

        //public static Dictionary<String, String> CopyTestReport(Dictionary<String, String> src_dest_file_list)
        //{
        //    Dictionary<String, String> Ret_copied_files = new Dictionary<String, String>();
        //    foreach (String src_file in src_dest_file_list.Keys)
        //    {
        //        if (Storage.FileExists(src_file))
        //        {
        //            String dest_file = src_dest_file_list[src_file];
        //            string dest_path = Storage.GetDirectoryName(dest_file);
        //            if (Storage.DirectoryExists(dest_path) == false)
        //            {
        //                Storage.CreateDirectory(dest_path, auto_parent_dir: true);
        //            }
        //            if (Storage.Copy(src_file, dest_file, overwrite: true))
        //            {
        //                Ret_copied_files.Add(src_file, dest_file);
        //            }
        //        }
        //        else
        //        {
        //            // file not exist
        //        }
        //    }

        //    return Ret_copied_files;
        //}

        // Copy and Clear judgement -- to be used for copying part/full of report from existing projects
        
        /*
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
                LogMessage.WriteLine("ERR: Open workbook in AutoCorrectReport_SingleFile(): " + source_file);
                return false;
            }

            String sheet_name = ReportGenerator.GetSheetNameAccordingToFilename(source_file);
            Worksheet ws = ExcelAction.Find_Worksheet(wb, sheet_name);

            //// If valid sheet_name does not exist, use first worksheet and rename it.
            //String sheet_name = ReportGenerator.GetSheetNameAccordingToFilename(source_file);
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
            KeywordReport.ClearJudgement(ws);

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
        */

    }
}
