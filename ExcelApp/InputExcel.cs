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

        /*
        public String Get_SRC_Directory()
        {
            String path = this.source_path;
            if (this.source_folder != "")
            {
                path = Storage.CominePath(path, this.source_folder);
            }
            if (this.source_group != "")
            {
                path = Storage.CominePath(path, this.source_group);
            }
            return path;
        }

        public String Get_DEST_Directory()
        {
            String path = this.destination_path;
            if (this.destination_folder != "")
            {
                path = Storage.CominePath(path, this.destination_folder);
            }
            if (this.destination_group != "")
            {
                path = Storage.CominePath(path, this.destination_group);
            }
            return path;
        }
        */

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

        static public String ExcelSheetName = "";

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
            Workbook wb = ExcelAction.OpenExcelWorkbook(filename: input_excel_file, ReadOnly: true);
            if (wb == null)
            {
                LogMessage.WriteLine("ERR: Open workbook in AutoCorrectReport_by_Excel(): " + input_excel_file);
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
            const int row_start_index  = 1;
            int row_index = row_start_index, col_index = 1;
            // start data processing since 2nd row
            row_index++;
            col_index = 1;
            List<CopyReport> report_list = new List<CopyReport>();
            do
            {
                CopyReport ctp = new CopyReport();
                ctp.source_path = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                //if (ctp.source_path != "")
                if (String.IsNullOrWhiteSpace(ctp.source_path) == false)
                {
                    ctp.source_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.source_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.source_report = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_path = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_report = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
                    ctp.destination_assignee = ExcelAction.GetCellTrimmedString(ws, row_index, col_index);
                    report_list.Add(ctp);
                    row_index++;
                    col_index = 1;
                }
                else
                {
                    bStillReadingExcel = false;
                    // If more than 1 row have data ==> get destination of 1st data row for return
                    if (row_index >= (row_start_index + 1))
                    {
                        return_destination_path = report_list[0].destination_path;
                    }
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
            foreach (CopyReport copy_report in report_list)
            {
                //String src_path = copy_report.Get_SRC_Directory();
                String src_fullfilename = copy_report.Get_SRC_FullFilePath();
                if (!Storage.FileExists(src_fullfilename))
                    continue;

                //String dest_path = copy_report.Get_DEST_Directory();
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
                    if (Storage.DirectoryExists(destination_dir) == false)
                    {
                        Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
                    }
                    success = Storage.Copy(source_file, destination_file, overwrite: true);
                }
                else // modifying contents so need to open excel
                {
                    String today = DateTime.Now.ToString("yyyy/MM/dd");
                    HeaderTemplate.UpdateVariables_TodayAssigneeLinkedIssue(today: today, assignee: assignee, LinkedIssue: StyleString.WhiteSpaceList());
                    success = KeywordReport.AutoCorrectReport_SingleFile(source_file: src, destination_file: dest, wb_template: wb, always_save: true);
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

            output_report_list = report_actually_copied_list_dest;

            if (report_cannot_be_copied_list_src.Count > 0)
                return false;   // some can't be copied
            else
                return true;

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
