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
            foreach (TestCase tc in ReportGenerator.ReadGlobalTestcaseList())
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

            if (report_actually_copied_list_src.Count > 0)
                return true;
            else
                return false;
        }

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

        static public bool AutoCorrectReport_SingleFile(String source_file, String destination_file, Workbook wb_template, Boolean always_save = false)
        {
           KeywordReportHeader out_header; // not used for this version of API.
           return AutoCorrectReport_SingleFile(source_file, destination_file, wb_template, out out_header, always_save);
        }
        
        // Copy and update worksheet name & header & bug-result (configurable) -- to be used for starting a new project based on reports from anywhere
        static public bool AutoCorrectReport_SingleFile(String source_file, String destination_file, Workbook wb_template, out KeywordReportHeader out_header, Boolean always_save = false)
        {
            Boolean file_has_been_updated = false;
            out_header = new KeywordReportHeader();

            destination_file = Storage.GetFullPath(destination_file);
            if (Storage.IsReportFilename(destination_file) == false)
            {
                // Do nothing if new filename does not look like a report filename.
                return file_has_been_updated;
            }

            // Open Excel workbook
            source_file = Storage.GetFullPath(source_file);
            Workbook wb = ExcelAction.OpenExcelWorkbook(filename: source_file, ReadOnly: false);
            if (wb == null)
            {
                ConsoleWarning("ERR: Open workbook in AutoCorrectReport_SingleFile(): " + source_file);
                return false;
            }

            // If valid sheet_name does not exist, use first worksheet .
            String current_sheet_name = TestPlan.GetSheetNameAccordingToFilename(source_file);
            KeywordReport.DefaultKeywordReportHeader.Report_SheetName = current_sheet_name;
            Worksheet ws;
            if (ExcelAction.WorksheetExist(wb, current_sheet_name) == false)
            {
                ws = wb.Sheets[1];
            }
            else
            {
                ws = ExcelAction.Find_Worksheet(wb, current_sheet_name);
            }

            // Update sheetname (when the option is true)
            if (KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Report_Sheetname)
            {
                String new_sheet_name = TestPlan.GetSheetNameAccordingToFilename(destination_file);
                ws.Name = new_sheet_name;
                KeywordReport.DefaultKeywordReportHeader.Report_SheetName = new_sheet_name;
                file_has_been_updated = true;
            }

            //Report_C_Update_Header_by_Template
            if (KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Header_by_Template == true)
            {
                if (ExcelAction.WorksheetExist(wb_template, HeaderTemplate.SheetName_HeaderTemplate))
                {
                    Worksheet ws_template = ExcelAction.Find_Worksheet(wb_template, HeaderTemplate.SheetName_HeaderTemplate);
                    String filename = TestPlan.GetReportTitleAccordingToFilename(destination_file);
                    String sheetname = ws.Name;
                    HeaderTemplate.UpdateVariables(filename: filename, sheetname: sheetname);
                    HeaderTemplate.CopyAndUpdateHeader(ws_template, ws);
                }
            }
            else
            {
                // Update header (when the option is true)
                if (KeywordReport.DefaultKeywordReportHeader.Report_C_Update_Full_Header == true)
                {
                    String new_title = TestPlan.GetReportTitleAccordingToFilename(destination_file);
                    KeywordReport.DefaultKeywordReportHeader.Report_Title = new_title;
                    // sheet-name is not defined as part of header --> it should be part of excel report (eg. filename, sheetname)
                    //KeywordReport.DefaultKeywordReportHeader.Report_SheetName = new_sheet_name;
                    KeywordReport.UpdateKeywordReportHeader_full(ws, KeywordReport.DefaultKeywordReportHeader);
                    file_has_been_updated = true;
                }

                //Report_C_Replace_Conclusion
                if (KeywordReport.DefaultKeywordReportHeader.Report_C_Replace_Conclusion == true)
                {
                    //StyleString blank_space = new StyleString(" ", StyleString.default_color, StyleString.default_font, StyleString.default_size);
                    StyleString blank_space = new StyleString(" ", StyleString.default_color, "Gill Sans MT", StyleString.default_size);
                    KeywordReport.ReplaceConclusionWithBugList(ws, blank_space.ConvertToList());
                    file_has_been_updated = true;
                }

                // Clear bug-list, bug-count, Pass/Fail/Conditional_Pass count, judgement
                if (KeywordReport.DefaultKeywordReportHeader.Report_C_Clear_Keyword_Result)
                {
                    KeywordReport.ClearKeywordBugResult(source_file, ws);
                    KeywordReport.ClearReportBugCount(ws);
                    KeywordReport.ClearJudgement(ws);
                    file_has_been_updated = true;
                }
            }

            // Hide keyword result/bug-list row -- after clear because it is un-hide after clear
            if (KeywordReport.DefaultKeywordReportHeader.Report_C_Hide_Keyword_Result_Bug_Row)
            {
                KeywordReport.HideKeywordResultBugRow(source_file, ws);
                file_has_been_updated = true;
            }

            if (KeywordReport.DefaultKeywordReportHeader.Report_C_Remove_AUO_Internal)
            {
                String sheet_name_to_keep = ws.Name;
                if (wb.Sheets.Count > 1)
                {
                    // work-sheet can be deleted only when there are two or more sheets
                    for (int sheet_index = wb.Sheets.Count; sheet_index > 0; sheet_index--)
                    {
                        String temp_sheet_name = wb.Sheets[sheet_index].Name;
                        if (temp_sheet_name.Length >= sheet_name_to_keep.Length)
                        {
                            if (temp_sheet_name.Substring(0, sheet_name_to_keep.Length) == sheet_name_to_keep)
                            {
                                continue;
                            }
                        }
                        wb.Sheets[sheet_index].Delete();
                        file_has_been_updated = true;
                    }
                }
            }

            if ((file_has_been_updated) || (always_save))
            {
                // Something has been updated or always save (ex: to copy file & update) ==> save to excel file
                String destination_dir = Storage.GetDirectoryName(destination_file);
                // if parent directory does not exist, create recursively all parents
                if (Storage.DirectoryExists(destination_dir) == false)
                {
                    Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
                }
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

        // Code for Report *.0
        public static int GroupSummary_Title_No_Row = 25, GroupSummary_Title_No_Col = 'D' - 'A' + 1;
        public static int GroupSummary_Title_TestItem_Row = 25, GroupSummary_Title_TestItem_Col = 'E' - 'A' + 1;
        public static int GroupSummary_Title_Result_Row = 25, GroupSummary_Title_Result_Col = 'H' - 'A' + 1;
        public static int GroupSummary_Title_Note_Row = 25, GroupSummary_Title_Note_Col = 'J' - 'A' + 1;
        public static String GroupSummary_Title_No_str = "No";
        public static String GroupSummary_Title_TestItem_str = "Test Item";
        public static String GroupSummary_Title_Result_str = "Result";
        public static String GroupSummary_Title_Note_str = "Note";

        static public Boolean Update_Single_Group_Summary_Report(Worksheet ws_group_report)
        {
            // check content of title row of group summary area, if not valid content, go to next
            if ((ExcelAction.CompareString(ws_group_report, GroupSummary_Title_No_Row, GroupSummary_Title_No_Col,GroupSummary_Title_No_str) == false ) ||
                (ExcelAction.CompareString(ws_group_report, GroupSummary_Title_TestItem_Row, GroupSummary_Title_TestItem_Col, GroupSummary_Title_TestItem_str) == false) ||
                (ExcelAction.CompareString(ws_group_report, GroupSummary_Title_Result_Row, GroupSummary_Title_Result_Col, GroupSummary_Title_Result_str) == false) ||
                (ExcelAction.CompareString(ws_group_report, GroupSummary_Title_Note_Row, GroupSummary_Title_Note_Col, GroupSummary_Title_Note_str) == false))
            {
                return false;
            }

            return true;
        }

        static public bool Update_Group_Summary(String report_path)
        {
            Boolean b_ret = false;
            String destination_path = Storage.GenerateDirectoryNameWithDateTime(report_path);

            // 1. List excel under report_path
            // 2. keep only "x.0" on the file list
            List<String> group_file_list = Storage.ListCandidateGroupSummaryFilesUnderDirectory(report_path);
            
            foreach (String group_file in group_file_list)
            {
                // 3. open
                // open standard test report
                Workbook wb_report = ExcelAction.OpenExcelWorkbook(group_file);
                if (wb_report == null)
                {
                    Console.WriteLine("OpenExcelWorkbook failed in Update_Group_Summary()");
                    continue;
                }

                String sheet_name = TestPlan.GetSheetNameAccordingToFilename(group_file);
                // Select and read work-sheet
                Worksheet ws_report = ExcelAction.Find_Worksheet(wb_report, sheet_name);
                if (ws_report == null)
                {
                    Console.WriteLine("Find_Worksheet (" + sheet_name + ") failed in Update_Group_Summary()");
                    ExcelAction.CloseExcelWorkbook(wb_report);
                    continue;
                }

                // 4. check content of title row, if not valid content, go to next
                if ((ExcelAction.GetCellTrimmedString(ws_report, GroupSummary_Title_No_Row, GroupSummary_Title_No_Col) != GroupSummary_Title_No_str) ||
                    (ExcelAction.GetCellTrimmedString(ws_report, GroupSummary_Title_TestItem_Row, GroupSummary_Title_TestItem_Col) != GroupSummary_Title_TestItem_str) ||
                    (ExcelAction.GetCellTrimmedString(ws_report, GroupSummary_Title_Result_Row, GroupSummary_Title_Result_Col) != GroupSummary_Title_Result_str) ||
                    (ExcelAction.GetCellTrimmedString(ws_report, GroupSummary_Title_Note_Row, GroupSummary_Title_Note_Col) != GroupSummary_Title_Note_str))
                {
                    ExcelAction.CloseExcelWorkbook(wb_report);
                    continue;
                }

                // 6. adjust table rows according to the number of TC summary within this gruop
                // count tc case in this group
                int tc_count = 0;
                if (tc_count == 0)
                {
                    ExcelAction.CloseExcelWorkbook(wb_report);
                    continue;
                }

                // find table lower boundary
                int row_index = GroupSummary_Title_No_Row + 1;
                int row_end = ExcelAction.Get_Range_RowNumber(ExcelAction.GetWorksheetAllRange(ws_report));
                Boolean row_found = false;
                while (!row_found && (row_index<=row_end))
                {
                    if (ExcelAction.GetCellTrimmedString(ws_report, row_index, GroupSummary_Title_No_Col) == "")
                    {
                        row_found = true;
                        row_index--;
                    }
                    else
                    {
                        row_index++;
                    }
                }

                // adjust row number

                
                // 7. Fill each row with (1) sheetname (2) summary after sheetname (3) TC result (4) linked issue on TC
                // for each tc item
                int current_row = 111;

                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_No_Col, "NO");
                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_TestItem_Col, "name");
                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_Result_Col, "judgement");
                ExcelAction.SetCellValue(ws_report, current_row, GroupSummary_Title_Note_Col, "linked issue");

                // 8. save and close and go to next report
                String output_filename = group_file.Replace(report_path, destination_path);
                ExcelAction.SaveExcelWorkbook(wb_report, output_filename);
                ExcelAction.CloseExcelWorkbook(wb_report);
            }

            b_ret = true;
            return b_ret;
        }

        // Code for copy-and-paste header
        static public Boolean CopyAndPasteHeaderTemplate()
        {
            Boolean b_ret = false;

            return b_ret;
        }

        // Code for updating header (according to special keyword)
        static public Boolean UpdatHeaderTemplate()
        {
            Boolean b_ret = false;

            return b_ret;
        }

        // Code for Report C
        static public bool AutoCorrectReport_by_Excel(String input_excel_file)
        {
            // open excel and read and close excel
            // Open Excel workbook
            Workbook wb = ExcelAction.OpenExcelWorkbook(filename: input_excel_file, ReadOnly: true);
            if (wb == null)
            {
                ConsoleWarning("ERR: Open workbook in AutoCorrectReport_by_Excel(): " + input_excel_file);
                return false;
            }

            Worksheet ws;
            if (ExcelAction.WorksheetExist(wb, HeaderTemplate.SheetName_ReportList))
            {
                ws = ExcelAction.Find_Worksheet(wb, HeaderTemplate.SheetName_ReportList);
            }
            else
            {
                ws = wb.ActiveSheet;
            }

            //public String source_path;
            //public String source_folder;
            //public String source_group;
            //public String source_filename;
            //public String destination_path;
            //public String destination_folder;
            //public String destination_group;
            //public String destination_filename;
            Boolean bStillReadingExcel = true;
            // check title row
            int row_index = 1, col_index = 1;
            // TBD
            row_index++;
            col_index = 1;
            List<CopyTestReport> report_list = new List<CopyTestReport>();
            do
            {
                CopyTestReport ctp = new CopyTestReport();
                ctp.source_path = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                //if (ctp.source_path != "")
                if (String.IsNullOrWhiteSpace(ctp.source_path)==false)
                {
                    ctp.source_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.source_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.source_filename = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_path = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_filename = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_assignee = ExcelAction.GetCellTrimmedString(ws, row_index, col_index);
                    report_list.Add(ctp);
                    row_index++;
                    col_index = 1;
                }
                else
                {
                    bStillReadingExcel = false;
                }
            }
            while (bStillReadingExcel);
            // Close later because excel is now also template for updating header so it will be used later
            //// Close Excel
            //ExcelAction.CloseExcelWorkbook(wb);

            // create list of source and destination
            Dictionary<String, String> copy_list = new Dictionary<String, String>();
            List<String> report_to_be_copied_list_src = new List<String>();
            List<String> report_to_be_copied_list_dest = new List<String>();
            List<String> report_to_be_copied_list_assignee = new List<String>();
            foreach (CopyTestReport copy_report in report_list)
            {
                String src_path = copy_report.Get_SRC_Directory();
                String src_fullfilename = copy_report.Get_SRC_FullFilePath();
                if (!Storage.FileExists(src_fullfilename))
                    continue;

                String dest_path = copy_report.Get_DEST_Directory();
                String dest_fullfilename = copy_report.Get_DEST_FullFilePath();
                report_to_be_copied_list_src.Add(src_fullfilename);
                report_to_be_copied_list_dest.Add(dest_fullfilename);
                report_to_be_copied_list_assignee.Add(copy_report.destination_assignee);
            }

            // auto-correct report files.
            List<String> report_actually_copied_list_src = new List<String>();
            List<String> report_actually_copied_list_dest = new List<String>();
            List<String> report_cannot_be_copied_list_src = new List<String>();
            List<String> report_cannot_be_copied_list_dest = new List<String>();
            // use Auto Correct Function to copy and auto-correct.

            for (int index = 0; index < report_to_be_copied_list_src.Count; index++)
            {
                String src = report_to_be_copied_list_src[index],
                       dest = report_to_be_copied_list_dest[index],
                       assignee = report_to_be_copied_list_assignee[index];
                Boolean success = false;

                // if only copying files, no need to open excel
                if (KeywordReport.DefaultKeywordReportHeader.Report_C_CopyFileOnly)
                {
                    String source_file = src, destination_file = dest;
                    String destination_dir = Storage.GetDirectoryName(destination_file);
                    // if parent directory does not exist, create recursively all parents
                    if (Storage.DirectoryExists(destination_dir)==false)
                    {
                        Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
                    }
                    success = Storage.Copy(source_file, destination_file, overwrite: true);
                }
                else // modifying contents so need to open excel
                {
                    String today = DateTime.Now.ToString("yyyy/MM/dd");
                    HeaderTemplate.UpdateVariables(today: today, assignee: assignee, LinkedIssue: StyleString.StringToListOfStyleString(" "));
                    success = AutoCorrectReport_SingleFile(source_file: src, destination_file: dest, wb_template: wb, always_save: true);
                }

                if (success)
                {
                    report_actually_copied_list_src.Add(src);
                    report_actually_copied_list_dest.Add(dest);
                }
                else
                {
                    report_cannot_be_copied_list_src.Add(src);
                    report_cannot_be_copied_list_dest.Add(dest);
                }
            }

            // Close Excel
            ExcelAction.CloseExcelWorkbook(wb);

            if (report_cannot_be_copied_list_src.Count > 0)
                return false;   // some can't be copied
            else
                return true;

        }

        static public bool AutoCorrectReport_by_Folder(String report_root, String Output_dir)
        {
            Boolean b_ret = false;

            // 0.1 List all files under report_root_dir.
            List<String> file_list = Storage.ListFilesUnderDirectory(report_root);
            // 0.2 filename check to exclude non-report files.
            List<String> report_filename = Storage.FilterFilename(file_list);

            foreach (String source_report in report_filename)
            {
                String dest_filename = KeywordReport.DecideDestinationFilename(report_root, Output_dir, source_report); // replace folder name
                b_ret |= AutoCorrectReport_SingleFile(source_file: source_report, destination_file: dest_filename, wb_template:new Workbook(), always_save: true);
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
