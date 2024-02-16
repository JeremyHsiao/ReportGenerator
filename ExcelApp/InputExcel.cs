using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ExcelReportApplication
{
    public class InputExcel
    {
        static public String SheetName_ReportList = "ReportList";
        static public String SheetName_HeaderTemplate_Source = "BeforeLine21";
        static public String SheetName_HeaderTemplate_Destination = "BeforeLine21";
        static public String SheetName_BugList = "Bug";
        static public String SheetName_TestCaseList = "TestCase";
        static public String SheetName_ReleaseNote = "ReleaseNote";
        static public String SheetName_ImportToJira = "ReportList";
        static public String SheetName_TestPlan = "TestPlan";
        static public String SheetName_AssigneeList = "AssigneeList";
        static public String SheetName_TCTemplate = "TCResult";

        static public void LoadFromXML()
        {
            SheetName_ReportList = XMLConfig.ReadAppSetting_String("Sheetname_ReportList");
            SheetName_HeaderTemplate_Source = XMLConfig.ReadAppSetting_String("Sheetname_HeaderTemplate_Source");
            SheetName_HeaderTemplate_Source = XMLConfig.ReadAppSetting_String("Sheetname_HeaderTemplate_Source");
            SheetName_HeaderTemplate_Destination = XMLConfig.ReadAppSetting_String("Sheetname_HeaderTemplate_Destination");
            SheetName_BugList = XMLConfig.ReadAppSetting_String("Sheetname_BugList");
            SheetName_TestCaseList = XMLConfig.ReadAppSetting_String("Sheetname_TestCaseList");
            SheetName_ReleaseNote = XMLConfig.ReadAppSetting_String("Sheetname_ReleaseNote");
            SheetName_ImportToJira = XMLConfig.ReadAppSetting_String("Sheetname_ImportToJira");
            SheetName_TestPlan = XMLConfig.ReadAppSetting_String("Sheetname_TestPlan");
            SheetName_AssigneeList = XMLConfig.ReadAppSetting_String("Sheetname_AssigneeList");
            SheetName_TCTemplate = XMLConfig.ReadAppSetting_String("Sheetname_TCTemplate");
        }

        public String inputExcelFilename;
        public Workbook workbook;
        public Worksheet activeWorksheet;
        public Worksheet reportListSheet;
        public Worksheet sourceTemplateSheet;
        public Worksheet destinationTemplateSheet;
        public Worksheet headerTemplateSheet;
        public Worksheet tcTemplateSheet;
        public Worksheet importToJiraSheet;

        public Boolean CheckSourceTemplateSheet()
        {
            sourceTemplateSheet = ExcelAction.Find_Worksheet(workbook, SheetName_HeaderTemplate_Source);
            return (sourceTemplateSheet != null) ? true : false;
        }

        public Boolean CheckDestinationTemplateSheet()
        {
            destinationTemplateSheet = ExcelAction.Find_Worksheet(workbook, SheetName_HeaderTemplate_Destination);
            return (destinationTemplateSheet != null) ? true : false;
        }

        public Boolean CheckHeaderTemplateSheet()
        {
            destinationTemplateSheet = ExcelAction.Find_Worksheet(workbook, SheetName_HeaderTemplate_Destination);
            return (destinationTemplateSheet != null) ? true : false;
        }

        public Boolean CheckReportListSheetActive()
        {
            if (ExcelAction.WorksheetExist(workbook, SheetName_ReportList))
            {
                workbook.Worksheets[SheetName_ReportList].Activate();
                activeWorksheet = reportListSheet = workbook.ActiveSheet;
            }
            return true;
        }

        public Boolean ProcessInputExcelHeaderTemplate(String input_excel_file)
        {
            inputExcelFilename = "";

            // Open Source Header Template Excel workbook
            workbook = ExcelAction.OpenExcelWorkbook(filename: input_excel_file, ReadOnly: true);
            if (workbook == null)
            {
                LogMessage.WriteLine("ERR: Open workbook failed in ProcessInputExcelHeaderTemplate(): " + input_excel_file);
                return false;
            }

            // Find source template sheet
            if (CheckSourceTemplateSheet() == false)
            {
                LogMessage.WriteLine("ERR: source template worksheet doesn't exist on excel: " + input_excel_file);
                return false;
            }

            // Find destination template sheet
            if (CheckDestinationTemplateSheet() == false)
            {
                LogMessage.WriteLine("ERR: destination template worksheet doesn't exist on excel: " + input_excel_file);
                return false;
            }

            if (CheckReportListSheetActive() == false)
            {
                LogMessage.WriteLine("ERR: report list worksheet doesn't exist on excel: " + input_excel_file);
                return false;
            }

            inputExcelFilename = input_excel_file;
            return true;
        }

        public Boolean ProcessInputExcelSelectReportList(String input_excel_file)
        {
            inputExcelFilename = "";

            // Open Source Header Template Excel workbook
            workbook = ExcelAction.OpenExcelWorkbook(filename: input_excel_file, ReadOnly: true);
            if (workbook == null)
            {
                LogMessage.WriteLine("ERR: Open workbook failed in ProcessInputExcelSelectReportList(): " + input_excel_file);
                return false;
            }

            if (CheckReportListSheetActive() == false)
            {
                LogMessage.WriteLine("ERR: report list worksheet doesn't exist on excel: " + input_excel_file);
                // fall-back solution for backward compatibility of old version input excel where active sheet is always report list independent of sheetname
                activeWorksheet = reportListSheet = workbook.ActiveSheet;
                // return false;
            }

            inputExcelFilename = input_excel_file;
            return true;
        }
    }

    public class CopyReport
    {
        private String source_path;
        private String source_folder;
        private String source_group;
        private String source_report;
        private String destination_path;
        private String destination_folder;
        private String destination_group;
        private String destination_report;
        private String destination_assignee;

        public String DestinationAssignee   // property
        {
            get { return destination_assignee; }   // get method
            set { destination_assignee = value; }  // set method
        }

        public String Get_SRC_FullFilePath()
        {
            String path = source_path;
            String folder = source_folder;
            String group = source_group;
            String report = source_report;

            String filename = report + ".xlsx";
            String filedir;
            Storage.CominePath(path, folder, group, out filedir);
            String fullfilepath = Storage.GetValidFullFilename(filedir, filename);
            return fullfilepath;
        }
        public String Get_DEST_FullFilePath()
        {
            String path = destination_path;
            String folder = destination_folder;
            String group = destination_group;
            String report = destination_report;

            String filename = report + ".xlsx";
            String filedir;
            Storage.CominePath(path, folder, group, out filedir);
            String fullfilepath = Storage.GetValidFullFilename(filedir, filename);
            return fullfilepath;
        }

        public Boolean ReadFromExcelRow(Worksheet worksheet, int row, int column)
        {
            Boolean b_ret = false;
            source_path = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
            if (String.IsNullOrWhiteSpace(source_path) == false)
            {
                source_folder = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
                source_group = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
                source_report = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
                destination_path = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
                destination_folder = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
                destination_group = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
                destination_report = ExcelAction.GetCellTrimmedString(worksheet, row, column++);
                destination_assignee = ExcelAction.GetCellTrimmedString(worksheet, row, column);
                b_ret = true;
            }
            return b_ret;
        }
        public Boolean WriteToExcelRow(Worksheet worksheet, int row, int column)
        {
            Boolean b_ret = false;
            ExcelAction.SetCellValue(worksheet, row, column++, source_path);
            ExcelAction.SetCellValue(worksheet, row, column++, source_folder);
            ExcelAction.SetCellValue(worksheet, row, column++, source_group);
            ExcelAction.SetCellValue(worksheet, row, column++, source_report);
            ExcelAction.SetCellValue(worksheet, row, column++, destination_path);
            ExcelAction.SetCellValue(worksheet, row, column++, destination_folder);
            ExcelAction.SetCellValue(worksheet, row, column++, destination_group);
            ExcelAction.SetCellValue(worksheet, row, column++, destination_report);
            ExcelAction.SetCellValue(worksheet, row, column++, destination_assignee);
            b_ret = true;
            return b_ret;
        }

        static private Boolean CreateEmptyErrorLogSheet(Workbook workbook, Worksheet source_worksheet, out Worksheet destination_worksheet)
        {
            Boolean b_ret = false;
            ExcelAction.DuplicateReportListSheet(source_worksheet);
            destination_worksheet = workbook.Sheets[source_worksheet.Index + 1];
            destination_worksheet.Rows["2:" + destination_worksheet.Rows.Count.ToString()].ClearContents();
            destination_worksheet.Name = source_worksheet.Name + "_ErrorLog";
            b_ret = true;
            return b_ret;
        }
        static public Boolean WriteErrorLog(Workbook workbook, Worksheet source_worksheet, List<CopyReport> source_inexist_list,
                                                        List<CopyReport> filename_not_report_list, List<CopyReport> process_fail_list)
        {
            Boolean b_ret = false;

            int err_log_row = 2, err_log_col = 1;
            Worksheet log_worksheet = source_worksheet;       // temporarily assignment
            String source_inexist_list_message = "Source Report Info contains some errors to be checked";
            String process_fail_list_message = "Some Report failed during processing -- to be checked";
            String filename_not_report_list_message = "Destination Info is not valid for report -- to be checked";
            String end_of_log_message = "End of Log";

            Boolean LogSheetNotYetCreated = true;

            if (source_inexist_list.Count > 0)
            {
                if (LogSheetNotYetCreated)      // not-yet failed --> need to copy a sheet for error log
                {
                    CreateEmptyErrorLogSheet(workbook, source_worksheet, out log_worksheet);
                    err_log_row = 2;
                    err_log_col = 1;
                    LogSheetNotYetCreated = false;
                }
                // write copy_fail_list
                ExcelAction.SetCellValue(log_worksheet, err_log_row, err_log_col, source_inexist_list_message);
                err_log_row++;
                err_log_col = 1;
                foreach (CopyReport cr in source_inexist_list)
                {
                    cr.WriteToExcelRow(log_worksheet, err_log_row, err_log_col);
                    err_log_row++;
                    err_log_col = 1;
                }
            }

            if (filename_not_report_list.Count > 0)
            {
                if (LogSheetNotYetCreated)      // not-yet failed --> need to copy a sheet for error log
                {
                    CreateEmptyErrorLogSheet(workbook, source_worksheet, out log_worksheet);
                    err_log_row = 2;
                    err_log_col = 1;
                    LogSheetNotYetCreated = false;
                }
                // write copy_fail_list
                ExcelAction.SetCellValue(log_worksheet, err_log_row, err_log_col, filename_not_report_list_message);
                err_log_row++;
                err_log_col = 1;
                foreach (CopyReport cr in filename_not_report_list)
                {
                    cr.WriteToExcelRow(log_worksheet, err_log_row, err_log_col);
                    err_log_row++;
                    err_log_col = 1;
                }
            }

            if (process_fail_list.Count > 0)
            {
                if (LogSheetNotYetCreated)      // not-yet failed --> need to copy a sheet for error log
                {
                    CreateEmptyErrorLogSheet(workbook, source_worksheet, out log_worksheet);
                    err_log_row = 2;
                    err_log_col = 1;
                    LogSheetNotYetCreated = false;
                }
                // write copy_fail_list
                ExcelAction.SetCellValue(log_worksheet, err_log_row, err_log_col, process_fail_list_message);
                err_log_row++;
                err_log_col = 1;
                foreach (CopyReport cr in process_fail_list)
                {
                    cr.WriteToExcelRow(log_worksheet, err_log_row, err_log_col);
                    err_log_row++;
                    err_log_col = 1;
                }
            }

            ExcelAction.CellActivate(log_worksheet, err_log_row, err_log_col);
            ExcelAction.SetCellValue(log_worksheet, err_log_row, err_log_col++, end_of_log_message);

            b_ret = true;
            return b_ret;
        }

        // Code for Report C
        static public Boolean UpdateTestReportByOptionAndSaveAsAnother(String input_excel_file)
        {
            List<String> report_list;
            String destination_path_1st_row;
            return UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(input_excel_file, out report_list, out destination_path_1st_row);
        }
        static public Boolean UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(String input_excel_file, out List<String> output_report_list, out String return_destination_path)
        {
            output_report_list = new List<String>();
            return_destination_path = "";

            InputExcel inputExcel = new InputExcel();

            if (inputExcel.ProcessInputExcelSelectReportList(input_excel_file) == false)
            {
                LogMessage.WriteLine("ERR: Failed in UpdateTestReportByOptionAndSaveAsAnother(): " + input_excel_file);
                return false;
            }

            Boolean bStillReadingExcel = true;
            // check title row
            const int row_start_index = 1;
            int row_index = row_start_index, col_index = 1;
            // start data processing since 2nd row
            row_index++;
            col_index = 1;
            List<CopyReport> report_list_to_be_processed = new List<CopyReport>();
            List<CopyReport> source_inexist_list = new List<CopyReport>();
            do
            {
                CopyReport ctp = new CopyReport();

                bStillReadingExcel = ctp.ReadFromExcelRow(inputExcel.reportListSheet, row_index, col_index);
                if (bStillReadingExcel)
                {
                    // Because copy-only doesn't need to check report filename condition, such check is done later not here
                    // Here only checking whether source_file is available.
                    String source_filename = ctp.Get_SRC_FullFilePath();
                    if (Storage.FileExists(source_filename))
                    {
                        report_list_to_be_processed.Add(ctp);
                    }
                    else
                    {
                        source_inexist_list.Add(ctp);
                    }
                    row_index++;
                    col_index = 1;
                }
            }
            while (bStillReadingExcel);

            List<CopyReport> destination_not_report_filename_list = new List<CopyReport>();
            List<CopyReport> process_success_list = new List<CopyReport>();
            List<CopyReport> process_fail_list = new List<CopyReport>();

            // if valid file-list, sort it (when required) before further processing
            if (report_list_to_be_processed.Count > 0)
            {
                if (TestReport.Option.FunctionC.CopyFileOnly)
                {
                    foreach (CopyReport cr in report_list_to_be_processed)
                    {
                        String source_file = cr.Get_SRC_FullFilePath();
                        String destination_file = cr.Get_DEST_FullFilePath();
                        String destination_dir = Storage.GetDirectoryName(destination_file);
                        Boolean success = false;
                        // if parent directory does not exist, create recursively all parents
                        if (Storage.DirectoryExists(destination_dir) == false)
                        {
                            Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
                        }

                        success = Storage.Copy(source_file, destination_file, overwrite: true);
                        if (success)
                        {
                            process_success_list.Add(cr);
                        }
                        else
                        {
                            process_fail_list.Add(cr);
                        }
                    }
                }
                else // (TestReport.Option.FunctionC.CopyFileOnly == false)
                {
                    // Sort in descending order of destination report sheetname (required for report processing with group summary report)
                    report_list_to_be_processed.Sort(CopyReport.Compare_by_Destination_Sheetname_Descending);

                    foreach (CopyReport cr in report_list_to_be_processed)
                    {
                        String source_file = cr.Get_SRC_FullFilePath();
                        String destination_file = cr.Get_DEST_FullFilePath();
                        String assignee = cr.destination_assignee;
                        Boolean success = false;

                        // only process when destination filename pass report/filename condition
                        if (Storage.IsReportFilename(destination_file))
                        {
                            String today = DateTime.Now.ToString("yyyy/MM/dd");
                            HeaderTemplate.UpdateVariables_TodayAssigneeLinkedIssue(today: today, assignee: assignee, LinkedIssue: StyleString.WhiteSpaceList());
                            // // if parent directory does not exist, FullyProcessReportSaveAsAnother() will create recursively all parents
                            success = TestReport.FullyProcessReportSaveAsAnother(source_file: source_file, destination_file: destination_file,
                                                        inputExcel: inputExcel, always_save: true);
                            if (success)
                            {
                                process_success_list.Add(cr);
                            }
                            else
                            {
                                process_fail_list.Add(cr);
                            }
                        }
                        else
                        {
                            destination_not_report_filename_list.Add(cr);
                        }
                    }
                }

                /*
                // Sort in descending order of destination report sheetname (required for report processing with group summary report)
                if (TestReport.Option.FunctionC.CopyFileOnly == false)
                {
                    report_list_to_be_processed.Sort(CopyReport.Compare_by_Destination_Sheetname_Descending);
                }

                foreach (CopyReport cr in report_list_to_be_processed)
                {
                    String src = cr.Get_SRC_FullFilePath(),
                           dest = cr.Get_DEST_FullFilePath(),
                           assignee = cr.destination_assignee;
                    Boolean success = false;

                    // if only copying files, no need to open excel
                    if (TestReport.Option.FunctionC.CopyFileOnly)
                    {
                        String source_file = src, destination_file = dest;
                        String destination_dir = Storage.GetDirectoryName(destination_file);
                        // if parent directory does not exist, create recursively all parents
                        if (Storage.DirectoryExists(destination_dir) == false)
                        {
                            Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
                        }
                        success = Storage.Copy(source_file, destination_file, overwrite: true);
                    }
                    else // modifying contents so need to open excel
                    {
                        String destination_filename = dest;
                        // only process when destination filename pass report/filename condition
                        if (Storage.IsReportFilename(destination_filename))
                        {
                            String today = DateTime.Now.ToString("yyyy/MM/dd");
                            HeaderTemplate.UpdateVariables_TodayAssigneeLinkedIssue(today: today, assignee: assignee, LinkedIssue: StyleString.WhiteSpaceList());
                            // // if parent directory does not exist, FullyProcessReportSaveAsAnother() will create recursively all parents
                            success = TestReport.FullyProcessReportSaveAsAnother(source_file: src, destination_file: dest, wb_header_template: wb_input_excel, always_save: true);
                        }
                    }

                    if (success)
                    {
                        process_success_list.Add(cr);
                    }
                    else
                    {
                        process_fail_list.Add(cr);
                    }
                }
                */
            }

            Boolean b_ret = true;
            if ((process_fail_list.Count > 0) || (destination_not_report_filename_list.Count > 0) || (source_inexist_list.Count > 0))
            {
                WriteErrorLog(inputExcel.workbook, inputExcel.reportListSheet, source_inexist_list, destination_not_report_filename_list, process_fail_list);
                b_ret = false;
                string new_filename = Storage.GenerateFilenameWithDateTime(inputExcel.inputExcelFilename);
                ExcelAction.CloseExcelWorkbook(workbook: inputExcel.workbook, SaveChanges: true, AsFilename: new_filename);
            }
            else
            {
                ExcelAction.CloseExcelWorkbook(inputExcel.workbook);
                b_ret = true;
            }

            return b_ret;
        }

        static public int Compare_by_Destination_Sheetname_Ascending(CopyReport report_x, CopyReport report_y)
        {

            String sheetname_x = ReportGenerator.GetSheetNameAccordingToFilename(report_x.Get_DEST_FullFilePath());
            String sheetname_y = ReportGenerator.GetSheetNameAccordingToFilename(report_y.Get_DEST_FullFilePath());

            return ReportGenerator.Compare_Sheetname_Ascending(sheetname_x, sheetname_y);
        }
        static public int Compare_by_Destination_Sheetname_Descending(CopyReport report_x, CopyReport report_y)
        {
            int compare_result_asceding = Compare_by_Destination_Sheetname_Ascending(report_x, report_y);
            return -compare_result_asceding;
        }

    }

    //public class ReportMapping
    //{
    //                //ctp.source_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.source_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.source_filename = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_path = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_filename = ExcelAction.GetCellTrimmedString(ws, row_index, col_index);
    //    private enum Name
    //    {
    //        source_path = 0,
    //        source_folder,
    //        source_group,
    //        source_filename,
    //        destination_path,
    //        destination_folder,
    //        destination_group,
    //        destination_filename,
    //        destination_assignee,
    //    }

    //    private static int EnumNameCount = Enum.GetNames(typeof(Name)).Length;

    //    private List<String> list_of_string;

    //    private void InitReportMapping()
    //    {
    //        list_of_string = new List<String>();
    //        for (int index = 0; index < EnumNameCount; index++)
    //        {
    //            list_of_string.Add("");
    //        }
    //    }

    //    // Note: here the sequence of member is pre-defined as source(path/folder/group/report), destination(path/folder/group/report/assignee)
    //    private void SetupReportMapping(List<String> member)
    //    {
    //        if (member.Count >= EnumNameCount)
    //        {
    //            int index = 0;
    //            do
    //            {
    //                list_of_string[index] = member[index];
    //            }
    //            while (++index < member.Count);
    //        }
    //        else
    //        {
    //            int index = 0;
    //            do
    //            {
    //                list_of_string[index] = member[index];
    //            }
    //            while (++index < member.Count);
    //            do
    //            {
    //                list_of_string[index] = "";
    //            }
    //            while (++index < EnumNameCount);
    //        }
    //    }

    //    public ReportMapping() { InitReportMapping(); }

    //    public ReportMapping(List<String> member) { InitReportMapping(); SetupReportMapping(member); }

    //    public String Source_Path   // property
    //    {
    //        get { return list_of_string[(int)Name.source_path]; }   // get method
    //        set { list_of_string[(int)Name.source_path] = value; }  // set method
    //    }

    //    public String Source_Folder   // property
    //    {
    //        get { return list_of_string[(int)Name.source_folder]; }   // get method
    //        set { list_of_string[(int)Name.source_folder] = value; }  // set method
    //    }

    //    public String Source_Group   // property
    //    {
    //        get { return list_of_string[(int)Name.source_group]; }   // get method
    //        set { list_of_string[(int)Name.source_group] = value; }  // set method
    //    }

    //    public String Source_Report   // property
    //    {
    //        get { return list_of_string[(int)Name.source_filename]; }   // get method
    //        set { list_of_string[(int)Name.source_filename] = value; }  // set method
    //    }

    //    public String Destination_Path   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_path]; }   // get method
    //        set { list_of_string[(int)Name.destination_path] = value; }  // set method
    //    }

    //    public String Destination_Folder   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_folder]; }   // get method
    //        set { list_of_string[(int)Name.destination_folder] = value; }  // set method
    //    }

    //    public String Destination_Group   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_group]; }   // get method
    //        set { list_of_string[(int)Name.destination_group] = value; }  // set method
    //    }

    //    public String Destination_Report   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_filename]; }   // get method
    //        set { list_of_string[(int)Name.destination_filename] = value; }  // set method
    //    }

    //    public String Destination_Assignee   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_assignee]; }   // get method
    //        set { list_of_string[(int)Name.destination_assignee] = value; }  // set method
    //    }

    //}

}
