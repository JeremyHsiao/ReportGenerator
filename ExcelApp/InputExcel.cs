using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ExcelReportApplication
{
    public class CopyReport
    {
        public String source_path;
        public String source_folder;
        public String source_group;
        public String source_report;
        public String destination_path;
        public String destination_folder;
        public String destination_group;
        public String destination_report;
        public String destination_assignee;

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

        // Code for Report C
        static public bool UpdateTestReportByOptionAndSaveAsAnother(String input_excel_file)
        {
            List<String> report_list;
            String destination_path_1st_row;
            return UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(input_excel_file, out report_list, out destination_path_1st_row);
        }
        static public bool UpdateTestReportByOptionAndSaveAsAnother_output_ReportList(String input_excel_file, out List<String> output_report_list, out String return_destination_path)
        {
            output_report_list = new List<String>();
            return_destination_path = "";

            // open excel and read and close excel
            // Open Excel workbook
            Workbook wb_input_excel = ExcelAction.OpenExcelWorkbook(filename: input_excel_file, ReadOnly: true);
            if (wb_input_excel == null)
            {
                LogMessage.WriteLine("ERR: Open workbook failed in UpdateTestReportByOptionAndSaveAsAnother(): " + input_excel_file);
                return false;
            }

            Worksheet ws_input_excel;
            if (ExcelAction.WorksheetExist(wb_input_excel, HeaderTemplate.SheetName_ReportList))
            {
                ws_input_excel = ExcelAction.Find_Worksheet(wb_input_excel, HeaderTemplate.SheetName_ReportList);
            }
            else
            {
                ws_input_excel = wb_input_excel.ActiveSheet;
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
                ctp.source_path = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                //if (ctp.source_path != "")
                if (String.IsNullOrWhiteSpace(ctp.source_path) == false)
                {
                    ctp.source_folder = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                    ctp.source_group = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                    ctp.source_report = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                    ctp.destination_path = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                    ctp.destination_folder = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                    ctp.destination_group = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                    ctp.destination_report = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index++);
                    ctp.destination_assignee = ExcelAction.GetCellTrimmedString(ws_input_excel, row_index, col_index);

                    // Because copy-only doesn't need to check report filename condition, such check is done later not here
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
                else
                {
                    bStillReadingExcel = false;
                }
            }
            while (bStillReadingExcel);
            // Close later because excel is now also template for updating header so it will be used later
            //// Close Excel
            //ExcelAction.CloseExcelWorkbook(wb);

            List<CopyReport> copy_success_list = new List<CopyReport>();
            List<CopyReport> copy_fail_list = new List<CopyReport>();

            // if valid file-list, sort it (when required) before further processing
            if (report_list_to_be_processed.Count > 0)
            {
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
                        copy_success_list.Add(cr);
                    }
                    else
                    {
                        copy_fail_list.Add(cr);
                    }
                }
            }

            Boolean b_ret = true;
            int err_log_row = 2, err_log_col = 1;
            Worksheet log_sheet = ws_input_excel;
            // If more than 1 row is ok to copy ==> get destination of 1st data row for out return
            if (copy_success_list.Count > 0)
            {
                return_destination_path = copy_success_list[0].destination_path;
                foreach (CopyReport cr in copy_success_list)
                {
                    output_report_list.Add(cr.Get_DEST_FullFilePath());
                }
            }

            if (source_inexist_list.Count > 0)
            {
                if (b_ret)      // not-yet failed --> need to copy a sheet for error log
                {
                    // copy excel and go to next one (just copied)
                    ExcelAction.DuplicateReportListSheet(ws_input_excel);
                    log_sheet = wb_input_excel.Sheets[ws_input_excel.Index + 1];
                    log_sheet.Rows["2:" + log_sheet.Rows.Count.ToString()].ClearContents();
                    b_ret = false;
                }
                // write copy_fail_list
                ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col, "Source Info to be checked");
                err_log_row++;
                err_log_col = 1;
                foreach (CopyReport cr in source_inexist_list)
                {
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_path);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_folder);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_group);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_report);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_path);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_folder);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_group);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_report);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_assignee);
                    err_log_row++;
                    err_log_col = 1;
                }
            }

            if (copy_fail_list.Count > 0)
            {
                if (b_ret)      // not-yet failed --> need to copy a sheet for error log
                {
                    // copy excel and go to next one (just copied)
                    ExcelAction.DuplicateReportListSheet(ws_input_excel);
                    log_sheet = wb_input_excel.Sheets[ws_input_excel.Index + 1];
                    log_sheet.Rows["2:" + log_sheet.Rows.Count.ToString()].ClearContents();
                    b_ret = false;
                }
                // write copy_fail_list
                ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col, "Items to be checked -- some error during processing");
                err_log_row++;
                err_log_col = 1;
                foreach (CopyReport cr in copy_fail_list)
                {
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_path);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_folder);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_group);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.source_report);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_path);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_folder);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_group);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_report);
                    ExcelAction.SetCellValue(log_sheet, err_log_row, err_log_col++, cr.destination_assignee);
                    err_log_row++;
                    err_log_col = 1;
                }
            }
            // Close Excel -- save if with error log
            if (b_ret)
            {
                ExcelAction.CloseExcelWorkbook(wb_input_excel);
            }
            else
            {
                string new_filename = Storage.GenerateFilenameWithDateTime(input_excel_file);
                ExcelAction.CloseExcelWorkbook(workbook: wb_input_excel, SaveChanges: true, AsFilename: new_filename);
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

    static public class HeaderTemplate
    {
        static public String SheetName_HeaderTemplate = "BeforeLine21";
        static public String SheetName_ReportList = "ReportList";

        static public int StartRow = 1;
        static public int EndRow = 9;
        static public int StartCol = 1;
        static public int EndCol = ExcelAction.ColumnNameToNumber('N');

        static public String Variable_ReportFileName = "$FileName$";
        static public String Variable_ReportSheetName = "$SheetName$";
        static public String Variable_Assignee = "$Assignee$";
        static public String Variable_Today = "$Today$";
        static public String Variable_TC_LinkedIssue = "$LinkedIssue$";
        static public String Variable_KEEP = "$KEEP$";

        static private String ReportFileName = "$FileName$";
        static private String ReportSheetName = "$SheetName$";
        static private String Assignee = "$Assignee$";
        static private String Today = "$Today$";
        static private List<StyleString> TC_LinkedIssue = StyleString.StringToListOfStyleString(Variable_TC_LinkedIssue);

        static public void ResetVariables()
        {
            ReportFileName = Variable_ReportFileName;
            ReportSheetName = Variable_ReportSheetName;
            Assignee = Variable_Assignee;
            Today = Variable_Today;
            TC_LinkedIssue = StyleString.StringToListOfStyleString(Variable_TC_LinkedIssue);
        }

        static public void UpdateVariables_FilenameSheetname(String filename, String sheetname)
        {
            ReportFileName = filename;
            ReportSheetName = sheetname;
        }

        static public void UpdateVariables_TodayAssigneeLinkedIssue(String today, String assignee, List<StyleString> LinkedIssue)
        {
            Assignee = assignee;
            Today = today;
            TC_LinkedIssue = LinkedIssue;
        }

        //static public void UpdateVariables(String filename = "", String sheetname = "", String assignee = "", String today = "", List<StyleString> LinkedIssue = null)
        //{
        //    if (String.IsNullOrWhiteSpace(filename) == false)
        //    {
        //        ReportFileName = filename;
        //    }
        //    if (String.IsNullOrWhiteSpace(sheetname) == false)
        //    {
        //        ReportSheetName = sheetname;
        //    }
        //    if (String.IsNullOrWhiteSpace(assignee) == false)
        //    {
        //        Assignee = assignee;
        //    }
        //    if (String.IsNullOrWhiteSpace(today) == false)
        //    {
        //        Today = today;
        //    }
        //    if (LinkedIssue != null)
        //    {
        //        TC_LinkedIssue = LinkedIssue;
        //    }
        //}

        static private Boolean CheckAndReplace(Worksheet ws, int row, int col, String from, String to)
        {
            Boolean b_ret = false;
            if (ExcelAction.GetCellValue(ws, row, col) != null)
            {
                String to_check = ExcelAction.GetCellValue(ws, row, col).ToString();
                if (to_check.Contains(from))
                {
                    ExcelAction.ReplaceText(ws, row, col, from, to);
                    b_ret = true;
                }
            }
            else
            {
                b_ret = true;
            }
            return b_ret;
        }

        static private Boolean CheckAndReplaceConclusion(Worksheet ws, int row, int col, String from, List<StyleString> to)
        {
            Boolean b_ret = false;
            if (ExcelAction.GetCellValue(ws, row, col) != null)
            {
                String to_check = ExcelAction.GetCellValue(ws, row, col).ToString();
                if (to_check.Contains(Variable_TC_LinkedIssue))
                {
                    StyleString.WriteStyleString(ws, row, col, TC_LinkedIssue);
                }
                b_ret = true;
            }
            else
            {
                b_ret = true;
            }
            return b_ret;
        }

        static public Boolean ReplaceHeaderVariableWithValue(Worksheet report_worksheet)
        {
            Boolean b_ret = false;

            for (int row_index = StartRow; row_index <= EndRow; row_index++)
            {
                for (int col_index = StartCol; col_index <= EndCol; col_index++)
                {
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_ReportFileName, ReportFileName);
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_ReportSheetName, ReportSheetName);
                    Assignee = Regex.Replace(Assignee, "[\u4E00-\u9FFF]", ""); // 移除中文
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_Assignee, Assignee);
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_Today, Today);
                    CheckAndReplaceConclusion(report_worksheet, row_index, col_index, Variable_TC_LinkedIssue, TC_LinkedIssue);
                }
            }
            b_ret = true;
            return b_ret;
        }

        static public Boolean CopyAndUpdateHeader(Worksheet template_worksheet, Worksheet report_worksheet)
        {
            Boolean b_ret = false;

            //ExcelAction.CopyRowHeight(template_worksheet, report_worksheet, StartRow, EndRow);
            //ExcelAction.CopyColumnWidth(template_worksheet, report_worksheet, StartCol, EndCol);
            ExcelAction.CopyPasteRows(template_worksheet, report_worksheet, StartRow, EndRow);

            b_ret = ReplaceHeaderVariableWithValue(report_worksheet);
            return b_ret;
        }

        static private List<int> KEEP_ROW = new List<int>(), KEEP_COL = new List<int>();
        static private List<Object> KEEP_CELL = new List<Object>();

        static public Boolean CopyKEEPCell(Worksheet template_worksheet, Worksheet report_workshee)
        {
            Boolean b_ret = false;
            KEEP_ROW.Clear();
            KEEP_COL.Clear();
            KEEP_CELL.Clear();
            for (int row_index = StartRow; row_index <= EndRow; row_index++)
            {
                for (int col_index = StartCol; col_index <= EndCol; col_index++)
                {
                    Object obj = ExcelAction.GetCellValue(template_worksheet, row_index, col_index);
                    if (obj != null)
                    {
                        String to_check = obj.ToString();
                        if (to_check.Contains(Variable_KEEP))
                        {
                            KEEP_ROW.Add(row_index);
                            KEEP_COL.Add(col_index);
                            obj = ExcelAction.GetCellValue(report_workshee, row_index, col_index);
                            KEEP_CELL.Add(obj);
                        }
                    }
                }
            }
            return b_ret;
        }

        static public Boolean CopyAndUpdateHeader_with_KEEP(Worksheet template_worksheet, Worksheet report_worksheet)
        {
            Boolean b_ret = false;

            CopyKEEPCell(template_worksheet, report_worksheet);
            //ExcelAction.CopyRowHeight(template_worksheet, report_worksheet, StartRow, EndRow);
            //ExcelAction.CopyColumnWidth(template_worksheet, report_worksheet, StartCol, EndCol);
            ExcelAction.CopyPasteRows(template_worksheet, report_worksheet, StartRow, EndRow);

            b_ret = ReplaceHeaderVariableWithValue(report_worksheet);

            // PasteKEEPCell
            int count = KEEP_ROW.Count();
            while (count-- > 0)
            {
                int row = KEEP_ROW[count], col = KEEP_COL[count];
                Object obj = KEEP_CELL[count];
                ExcelAction.SetCellValue(report_worksheet, row, col, obj);
            }

            return b_ret;
        }

    }

}
