using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Drawing;

namespace ExcelReportApplication
{
    static public class ExcelAction
    {
        static private Excel.Application excel_app;
        static private Workbook workbook_issuelist;
        static private Worksheet ws_issuelist;
        static private Workbook workbook_testcase;
        static private Worksheet ws_testcase;
        static private Workbook workbook_tc_template;
        static private Worksheet ws_tc_template;
        //static private Workbook workbook_testplan;
        //static private Worksheet ws_testplan;
        //static private Workbook workbook_keywordlog_template;
        static private Worksheet ws_keyword_list;
        static private Worksheet ws_not_keyword_file;
        static private Workbook workbook_new_keyword_list;
        static public List<String> Error_Log = new List<String>();

        static public bool ExcelVisible = true;

        static public Boolean OpenExcelApp()
        {
            Boolean b_ret = false;

            if (excel_app != null) return true;  //  if not null --> it can be used?!
            try
            {
                excel_app = new Excel.Application();
                excel_app.Visible = ExcelVisible;
                excel_app.Caption = "DQA Report Generator";
                excel_app.DisplayAlerts = false;
                b_ret = true;
            }
            catch (Exception ex)
            {
                LogMessage.WriteLine("OpenExcelApp has expection");
            }
            return b_ret;
        }

        static public Workbook OpenExcelWorkbook(String filename, bool ReadOnly = true, bool XLS = false, bool UpdateLinks = false)
        {
            Workbook ret_workbook = null;
            Boolean try_openXML = false;

            // Protection-check: file must exist.
            if (!Storage.FileExists(filename))
            {
                return ret_workbook;
            }

            try
            {
                if (XLS)
                {
                    ret_workbook = excel_app.Workbooks.Open(filename, ReadOnly: ReadOnly, CorruptLoad: XlCorruptLoad.xlExtractData,
                                                            UpdateLinks: UpdateLinks);
                }
                else
                {
                    ret_workbook = excel_app.Workbooks.Open(filename, ReadOnly: ReadOnly,
                                                            UpdateLinks: UpdateLinks);
                }
            }
            catch (Exception ex)
            {
                // if exception, try
                try_openXML = true;
            }

            if (!try_openXML)
            {
                return ret_workbook;
            }

            try
            {
                ret_workbook = excel_app.Workbooks.OpenXML(Filename: filename, LoadOption: XlXmlLoadOption.xlXmlLoadOpenXml);
            }
            catch (Exception ex)
            {
                Error_Log.Add("OpenExcelWorkbook exception: " + filename); // here is for setting break-point
            }

            return ret_workbook;
        }

        static public void SaveExcelWorkbook(Workbook workbook, String filename)
        {
            if (String.IsNullOrWhiteSpace(filename) == false)
            {
                // Try suggestion of https://stackoverflow.com/questions/5188527/how-to-deal-with-files-with-a-name-longer-than-259-characters
                String retry_filename = @"\\?\" + filename, final_filename = " ";

                try
                {
                    if (filename.Length < 259)
                    {
                        workbook.SaveAs(filename);
                        final_filename = filename;
                    }
                    else
                    {
                        workbook.SaveAs(retry_filename);
                        final_filename = retry_filename;
                    }
                }
                catch (Exception ex)
                {
                    Error_Log.Add("SaveExcelWorkbook exception: " + final_filename); // here is for setting break-point
                }
            }
        }

        static public void CloseExcelWorkbook(Workbook workbook, bool SaveChanges = false, String AsFilename = "")
        {
            excel_app.DisplayAlerts = false;
            if (SaveChanges)
            {
                if (String.IsNullOrWhiteSpace(AsFilename) == false)
                {
                    // Try suggestion of https://stackoverflow.com/questions/5188527/how-to-deal-with-files-with-a-name-longer-than-259-characters
                    String retry_filename = @"\\?\" + AsFilename, final_filename = " ";

                    try
                    {
                        if (AsFilename.Length < 259)
                        {
                            workbook.Close(SaveChanges: true, Filename: AsFilename);
                            final_filename = AsFilename;
                        }
                        else
                        {
                            workbook.Close(SaveChanges: true, Filename: retry_filename);
                            final_filename = retry_filename;
                        }
                    }
                    catch (Exception ex)
                    {
                        Error_Log.Add("CloseExcelWorkbook exception: " + final_filename); // here is for setting break-point
                    }
                }
                else
                {
                    workbook.Close(SaveChanges: true);
                }
            }
            else
            {
                workbook.Close(SaveChanges: false);
            }
        }

        static public void CloseExcelApp()
        {
            if (excel_app == null) return;
            excel_app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app);
            GC.Collect();
            excel_app = null;
        }

        // List all worksheets within excel 
        static public List<String> ListSheetName(Excel.Application curExcel)
        {
            List<String> ret_sheetname = new List<String>();
            foreach (Worksheet displayWorksheet in curExcel.Worksheets)
            {
                ret_sheetname.Add(displayWorksheet.Name);
            }
            return ret_sheetname;
        }

        static public bool WorksheetExist(Excel.Application curExcel, string sheet_name)
        {
            foreach (Worksheet displayWorksheet in curExcel.Worksheets)
            {
                if (displayWorksheet.Name.CompareTo(sheet_name) == 0)
                {
                    return true;
                }
            }
            return false;
        }

        static public bool WorksheetExist(Workbook wb, string sheet_name)
        {
            foreach (Worksheet displayWorksheet in wb.Worksheets)
            {
                if (displayWorksheet.Name.CompareTo(sheet_name) == 0)
                {
                    return true;
                }
            }
            return false;
        }

        // return worksheet with specified sheet_name; return null if not found
        static public Worksheet Find_Worksheet(Excel.Application curExcel, string sheet_name)
        {
            Worksheet ret = null;
            if (WorksheetExist(curExcel, sheet_name))
            {
                ret = curExcel.Sheets.Item[sheet_name];
            }
            return ret;
        }

        static public Worksheet Find_Worksheet(Workbook wb, string sheet_name)
        {
            Worksheet ret = null;
            if (WorksheetExist(wb, sheet_name))
            {
                ret = wb.Worksheets[sheet_name];
            }
            return ret;
        }

        static public Range GetWorksheetAllRange(Worksheet ws)
        {
            // For XlSpecialCellsValue,
            // please refer to https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlspecialcellsvalue?view=excel-pia
            Range ret_range = ws.Range["A1"].SpecialCells(XlCellType.xlCellTypeLastCell, (XlSpecialCellsValue)(1 + 2 + 4 + 16));
            return ret_range;
        }

        static public int Get_Range_RowNumber(Range check_range)
        {
            return (check_range.Row * check_range.Rows.Count);
        }

        static public int Get_Range_ColumnNumber(Range check_range)
        {
            return (check_range.Column * check_range.Columns.Count);
        }

        static public Range GetWorksheetPrintableRange(Worksheet ws)
        {
            String PrintArea = ws.PageSetup.PrintArea;
            Range rngPrintable;
            try
            {
                rngPrintable = ws.Range[PrintArea];
            }
            catch
            {
                // Use whole sheet as workaround for Printable Range
                rngPrintable = ws.Range["A1"].SpecialCells(XlCellType.xlCellTypeLastCell, (XlSpecialCellsValue)(1 + 2 + 4 + 16));
                Error_Log.Add("GetWorksheetPrintableRange exception with PrintArea: " + PrintArea); // here is for setting break-point
            }
            return rngPrintable;
        }

        static public void AutoFit_Column(Worksheet ws, int col)
        {
            ws.Columns[col].AutoFit();
        }

        static public void AutoFit_Row(Worksheet ws, int row)
        {
            ws.Rows[row].AutoFit();
        }

        static public double Get_Column_Width(Worksheet ws, int col)
        {
            double col_width;
            try
            {
                col_width = ws.Columns[col].ColumnWidth;
            }
            catch
            {
                // TBD: need to replace this workaround with better solution.
                col_width = ws.Columns[1].ColumnWidth;
                Error_Log.Add("Get_Column_Width exception with col_width: " + col_width.ToString()); // here is for setting break-point
            }
            return col_width;
        }

        static public void Set_Column_Width(Worksheet ws, int col, double width)
        {
            try
            {
                ws.Rows[col].ColumnWidth = width;
            }
            catch
            {
                // TBD: deal with exception when RowHeight can't be set via ws.Rows[row].RowHeight = height;
                Error_Log.Add("Set_Column_Width exception with width: " + width.ToString()); // here is for setting break-point
            }
        }

        static public void Set_Row_Height(Worksheet ws, int row, double height)
        {
            try
            {
                ws.Rows[row].RowHeight = height;
            }
            catch
            {
                // TBD: deal with exception when RowHeight can't be set via ws.Rows[row].RowHeight = height;
                Error_Log.Add("Set_Row_Height exception with height: " + height.ToString()); // here is for setting break-point
            }
        }

        static public double Get_Row_Height(Worksheet ws, int row)
        {
            double row_height;
            try
            {
                row_height = ws.Rows[row].RowHeight;
            }
            catch
            {
                // TBD: need to replace this workaround with better solution.
                row_height = ws.Rows[1].RowHeight;
                Error_Log.Add("Get_Row_Height exception with height: " + row_height.ToString()); // here is for setting break-point
            }
            return row_height;
        }

        static public void Unhide_Row(Worksheet ws, int row, int count = 1)
        {
            Range hiddenRange = ws.Range[ws.Cells[row, 1], ws.Cells[row + count - 1, Get_Range_ColumnNumber(GetWorksheetAllRange(ws))]];
            //     var hiddenRange = yourWorksheet.Range[yourWorksheet.Cells[firstRowToHide, firstColToHide], yourWorksheet.Cells[lastRowToHide, lastColToHide]];
            hiddenRange.EntireRow.Hidden = false;
        }

        static public void Hide_Row(Worksheet ws, int row, int count = 1)
        {
            Range hiddenRange = ws.Range[ws.Cells[row, 1], ws.Cells[row + count - 1, Get_Range_ColumnNumber(GetWorksheetAllRange(ws))]];
            //     var hiddenRange = yourWorksheet.Range[yourWorksheet.Cells[firstRowToHide, firstColToHide], yourWorksheet.Cells[lastRowToHide, lastColToHide]];
            hiddenRange.EntireRow.Hidden = true;
        }

        static public void CopyColumnWidth(Worksheet source_ws, Worksheet destination_ws, int StartCol, int EndCol)
        {
            for (int index = StartCol; index <= EndCol; index++)
            {
                destination_ws.Columns[index].ColumnWidth = source_ws.Columns[index].ColumnWidth;
            }
        }

        static public void CopyRowHeight(Worksheet source_ws, Worksheet destination_ws, int StartRow, int EndRow)
        {
            for (int index = StartRow; index <= EndRow; index++)
            {
                destination_ws.Rows[index].RowHeight = source_ws.Rows[index].RowHeight;
            }
        }

        static public void Insert_Column(Worksheet ws, int at_col)
        {
            ws.Columns[at_col].Insert();
        }

        static public void Insert_Row(Worksheet ws, int at_row)
        {
            ws.Rows[at_row].Insert();
        }

        static public void Delete_Row(Worksheet ws, int at_row)
        {
            ws.Rows[at_row].Delete();
        }

        //static public void SetupHorizontalLinkBreakToColumnH(Worksheet ws)
        //{
        //    if(ws.HPageBreaks[0] = 
        //}

        static public void CellTextAlignLeft(Worksheet ws, int at_row, int at_col)
        {
            ws.Cells[at_row, at_col].HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        static public void CellTextAlignUpperLeft(Worksheet ws, int at_row, int at_col)
        {
            ws.Cells[at_row, at_col].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            ws.Cells[at_row, at_col].VerticalAlignment = XlVAlign.xlVAlignTop;
        }

        static public void CellActivate(Worksheet ws, int at_row, int at_col)
        {
            //disable for fewer error caused by user
            //ws.Cells[at_row, at_col].Select();
        }

        static public Dictionary<string, int> CreateTableColumnIndex(Worksheet ws, int title_row)
        {
            Dictionary<string, int> col_name_list = new Dictionary<string, int>();

            int col_end = Get_Range_ColumnNumber(GetWorksheetAllRange(ws));
            for (int col_index = 1; col_index <= col_end; col_index++)
            {
                String cell_value2 = GetCellTrimmedString(ws, title_row, col_index);
                //if (cell_value2 == "") { continue; }
                if (String.IsNullOrWhiteSpace(cell_value2))
                {
                    LogMessage.WriteLine("TC/Template title is empty at column: " + col_index);
                    continue;               // column header is empty. shouldn't be here
                }

                if (col_name_list.ContainsKey(cell_value2))
                {
                    LogMessage.WriteLine("TC/Template title is duplicated at column: " + col_index);
                    continue;               // column header is repeated. shouldn't be here
                }
                col_name_list.Add(cell_value2.ToString(), col_index);
            }

            return col_name_list;
        }

        // Code for operations on specific Excel File

        public enum ExcelStatus
        {
            OK = 0,
            INIT_STATE,
            ERR_OpenIssueListExcel_OpenExcelWorkbook,
            ERR_OpenIssueListExcel_Find_Worksheet,
            ERR_OpenTestCaseExcel_OpenExcelWorkbook,
            ERR_OpenTestCaseExcel_Find_Worksheet,
            ERR_OpenKeywordLogTemplateExcel_OpenExcelWorkbook,
            ERR_OpenKeywordLogTemplateExcel_Find_Keyword_Worksheet,
            ERR_OpenKeywordLogTemplateExcel_Find_NonKeyword_Worksheet,
            ERR_CloseIssueListExcel_wb_null,
            ERR_CloseTestCaseExcel_wb_null,
            ERR_CloseKeywordLogTemplateExcel_wb_null,
            ERR_SaveChangesAndCloseIssueListExcel_wb_null,
            ERR_SaveChangesAndCloseTestCaseExcel_wb_null,
            ERR_SaveChangesAndCloseKeywordLogTemplateExcel_wb_null,
            ERR_NOT_DEFINED,
            EX_OpenIssueListWorksheet,
            EX_CloseIssueListExcel,
            EX_SaveChangesAndCloseIssueListExcel,
            EX_OpenTestCaseWorksheet,
            EX_CloseTestCaseWorksheet,
            EX_SaveChangesAndCloseTestCaseExcel,
            EX_OpenKeywordLogTemplateWorksheet,
            EX_CloseKeywordLogTemplateExcel,
            EX_SaveChangesAndCloseKeywordLogTemplateExcel,
            MAX_NO
        };

        // Excel accessing function for Issue List Excel

        static public Worksheet GetIssueListWorksheet()
        {
            return ws_issuelist;
        }

        static public Range GetIssueListAllRange()
        {
            return GetWorksheetAllRange(ws_issuelist);
        }

        //static public int GetIssueListMaxRow()
        //{
        //int max_row = GetWorksheetAllRange(ws_issuelist).Row;
        //if(max_row
        //}

        static public Object GetIssueListCell(int row, int col)
        {
            return ws_issuelist.Cells[row, col].Value2;
        }

        static public String GetIssueListCellTrimmedString(int row, int col)
        {
            Object cell_value2 = GetIssueListCell(row, col);
            if (cell_value2 == null) { return ""; }
            return cell_value2.ToString();
        }

        static public void IssueList_AutoFit_Column(int col)
        {
            AutoFit_Column(ws_issuelist, col);
        }

        static public void IssueList_WriteStyleString(int row, int col, List<StyleString> style_string_list)
        {
            StyleString.WriteStyleString(ws_issuelist, row, col, style_string_list);
        }

        static public Dictionary<string, int> CreateIssueListColumnIndex(int title_row_no)
        {
            return CreateTableColumnIndex(ws_issuelist, title_row_no);
        }

        // Excel accessing function for Test Case Excel

        static public Worksheet GetTestCaseWorksheet(bool IsTemplate = false)
        {
            return ((IsTemplate) ? ws_tc_template : ws_testcase);
        }

        static public Range GetTestCaseAllRange(bool IsTemplate = false)
        {
            return GetWorksheetAllRange(((IsTemplate) ? ws_tc_template : ws_testcase));
        }

        static public Object GetTestCaseCell(int row, int col, bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            return ws.Cells[row, col].Value2;
        }

        static public void SetTestCaseCell(int row, int col, Object set_object, bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            ws.Cells[row, col].Value2 = set_object;
        }

        static public String GetTestCaseCellTrimmedString(int row, int col, bool IsTemplate = false)
        {
            Object cell_value2 = GetTestCaseCell(row, col, IsTemplate: IsTemplate);
            if (cell_value2 == null) { return ""; }
            return cell_value2.ToString().Trim();
        }

        static public void ReplaceText(Worksheet ws, int row, int col, String from, String to)
        {
            Range cell = ws.Cells[row, col];
            String cell_text = ws.Cells[row, col].Value2.ToString();
            String new_cell_text = cell_text.Replace(from, to);
            cell.Value2 = new_cell_text;
        }

        //static public void SetKeywordListCell(int row, int col, Object set_object, Boolean to_keyword_list = true)
        //{
        //    Worksheet ws = ((to_keyword_list) ? ws_keyword_list : ws_not_keyword_file);
        //    ws.Cells[row, col].Value2 = set_object;
        //}

        static public Object GetCellValue(Worksheet ws, int row, int col)
        {
            return ws.Cells[row, col].Value2;
        }

        static public void SetCellValue(Worksheet ws, int row, int col, Object value, Boolean ClearContentFirst = false)
        {
            Range cell = ws.Cells[row, col];
            if (ClearContentFirst)
            {
                cell.MergeArea.ClearContents();
            }
            cell.NumberFormat = "@";
            cell.Value2 = value;
        }

        static public void SetCellString(Worksheet ws, int row, int col, String value)
        {
            Range cell = ws.Cells[row, col];
            cell.Value2 = value;
        }

        static public String GetCellTrimmedString(Worksheet ws, int row, int col)
        {
            Object cell_value2 = GetCellValue(ws, row, col);
            if (cell_value2 == null) { return ""; }
            return cell_value2.ToString().Trim();
        }

        static public Boolean CompareString(Worksheet ws, int row, int col, String to_compare)
        {
            String temp = GetCellTrimmedString(ws, row, col);
            return (temp == to_compare);
        }

        static public void TestCase_CellActivate(int at_row, int at_col, bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            CellActivate(ws, at_row, at_col);
        }

        static public void TestCase_AutoFit_Column(int col, bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            AutoFit_Column(ws, col);
        }

        static public void TestCase_Hide_Row(int row, int count = 1, bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            Hide_Row(ws, row, count);
        }

        static public void TestCase_WriteStyleString(int row, int col, List<StyleString> style_string_list, bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            StyleString.WriteStyleString(ws, row, col, style_string_list);
        }

        /*
                static public void CopyTestCaseIntoTemplate(String template_filename)
                {
                    ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(template_filename, IsTemplate: true);
                    if (status == ExcelAction.ExcelStatus.OK)
                    {
                        Range Src = GetWorksheetAllRange(ws_testcase);
                        Range Dst = GetWorksheetAllRange(ws_tc_template);
                        int Src_last_row = Src.Row, Src_last_col = Src.Column;
                        int Dst_last_row = Dst.Row, Dst_last_col = Dst.Column;

                        // Make template row count == TestCase row count
                        if (Src_last_row > Dst_last_row)
                        {
                            // Insert row into template file
                            int rows_to_insert = Src_last_row - Dst_last_row;
                            do
                            {
                                ws_tc_template.Rows[TestCase.DataBeginRow].Insert();
                            }
                            while (--rows_to_insert > 0);
                        }
                        else if (Src_last_row < Dst_last_row)
                        {
                            // Delete row from template file
                            int rows_to_delete = Dst_last_row - Src_last_row;
                            do
                            {
                                ws_tc_template.Rows[TestCase.DataBeginRow].Delete();
                            }
                            while (--rows_to_delete > 0);
                        }

                        // Copy Value from row (TestCase.DataBeginRow) to last-1 
                        Src = ws_testcase.Range[ws_testcase.Cells[TestCase.DataBeginRow, 1], ws_testcase.Cells[Src_last_row - 1, Src_last_col]];
                        Dst = ws_tc_template.Range[ws_tc_template.Cells[TestCase.DataBeginRow, 1], ws_tc_template.Cells[Src_last_row - 1, Src_last_col]];
                        Src.Copy(Type.Missing);
                        Dst.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

         -`-               // Format row 1-4 (TestCase.DataBeginRow-1)
                        //Src = ws_tc_template.Range[ws_tc_template.Cells[1, 1], ws_tc_template.Cells[TestCase.DataBeginRow - 1, Src_last_col]];
                        //Dst = ws_testcase.Range[ws_testcase.Cells[1, 1], ws_testcase.Cells[TestCase.DataBeginRow-1, Dst_last_col]];
                        //Src.Copy(Type.Missing);
                        //Dst.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        // Format 5 (TestCase.DataBeginRow) to last-1 
                        //Src = ws_tc_template.Range[ws_tc_template.Cells[TestCase.DataBeginRow, 1], ws_tc_template.Cells[TestCase.DataBeginRow, Src_last_col]];
                        //Dst = ws_testcase.Range[ws_testcase.Cells[TestCase.DataBeginRow, 1], ws_testcase.Cells[TestCase.DataBeginRow, Dst_last_col]];
                        //Src.Copy(Type.Missing);
                        //Dst.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        // Format last
                        //Src = ws_tc_template.Range[ws_tc_template.Cells[Src_last_row, 1], ws_tc_template.Cells[Src_last_row, Src_last_col]];
                        //Dst = ws_testcase.Range[ws_testcase.Cells[Dst_last_row, 1], ws_testcase.Cells[Dst_last_row - 1, Dst_last_col]];
                        //Src.Copy(Type.Missing);
                        //Dst.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                

                        ExcelAction.CloseTestCaseExcel(IsTemplate: true);
                    }
                }
        */

        static public Dictionary<string, int> CreateTestCaseColumnIndex(int title_row_no, Boolean IsTemplate)
        {
            if (IsTemplate)
            {
                return CreateTableColumnIndex(ws_tc_template, title_row_no);
            }
            else
            {
                return CreateTableColumnIndex(ws_testcase, title_row_no);
            }
        }

        static public int GetTestCaseExcelRange_Col(bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            int col_end = Get_Range_ColumnNumber(GetWorksheetAllRange(ws));
            return col_end;
        }

/*
        static public Dictionary<string, int> CreateBugListColumnIndex()
        {
            return CreateTableColumnIndex(ws_issuelist, Issue.NameDefinitionRow);
        }
*/

        static public int GetBugListExcelRange_Col()
        {
            Worksheet ws = (ws_issuelist);
            int col_end = Get_Range_ColumnNumber(GetWorksheetAllRange(ws));
            return col_end;
        }

        // Excel Open/Close/Save for Issue List Excel

        static public ExcelStatus OpenIssueListExcel(String buglist_filename)
        {
            try
            {
                Workbook wb_issuelist;

                if (Storage.GetExtension(buglist_filename) == ".xls")
                {
                    // Open excel (read-only & corrupt-load)
                    wb_issuelist = ExcelAction.OpenExcelWorkbook(buglist_filename, XLS: true);
                    MainForm.SystemLogAddLine("Bug file is XML");
                }
                else
                {
                    wb_issuelist = ExcelAction.OpenExcelWorkbook(buglist_filename);
                    MainForm.SystemLogAddLine("Bug file is XLS(X)");
                }

                if (wb_issuelist == null)
                {
                    return ExcelStatus.ERR_OpenIssueListExcel_OpenExcelWorkbook;
                }

                Worksheet ws_buglist = ExcelAction.Find_Worksheet(wb_issuelist, Issue.SheetName);
                if (ws_buglist == null)
                {
                    MainForm.SystemLogAddLine("Bug file does not have specified sheetname");
                    return ExcelStatus.ERR_OpenIssueListExcel_Find_Worksheet;
                }
                else
                {
                    workbook_issuelist = wb_issuelist;
                    ws_issuelist = ws_buglist;
                    return ExcelStatus.OK;
                }
            }
            catch
            {
                MainForm.SystemLogAddLine("OpenIssueListExcel has exception");
                return ExcelStatus.EX_OpenIssueListWorksheet;
            }

            // Not needed because never reaching here
            //return ExcelStatus.ERR_NOT_DEFINED;
        }

        static public ExcelStatus CloseIssueListExcel()
        {
            try
            {
                if (workbook_issuelist == null)
                {
                    return ExcelStatus.ERR_CloseIssueListExcel_wb_null;
                }
                ExcelAction.CloseExcelWorkbook(workbook_issuelist, SaveChanges: false);
                ws_issuelist = null;
                workbook_issuelist = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_issuelist = null;
                workbook_issuelist = null;
                return ExcelStatus.EX_CloseIssueListExcel;
            }
        }

        static public ExcelStatus SaveChangesAndCloseIssueListExcel(String dest_filename)
        {
            try
            {
                if (workbook_issuelist == null)
                {
                    return ExcelStatus.ERR_SaveChangesAndCloseIssueListExcel_wb_null;
                }
                ExcelAction.CloseExcelWorkbook(workbook_issuelist, SaveChanges: true, AsFilename: dest_filename);
                ws_issuelist = null;
                workbook_issuelist = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_issuelist = null;
                workbook_issuelist = null;
                return ExcelStatus.EX_SaveChangesAndCloseIssueListExcel;
            }
            // Not needed because never reaching here
            //return ExcelStatus.ERR_NOT_DEFINED;
        }

        // Excel Open/Close/Save for Test Case Excel

        static public ExcelStatus OpenTestCaseExcel(String tclist_filename, bool IsTemplate = false)
        {
            try
            {
                Workbook wb_tc;
                if (IsTemplate == false)
                {
                    if (Storage.GetExtension(tclist_filename) == ".xls")
                    {
                        // Open excel (read-only & corrupt-load)
                        wb_tc = ExcelAction.OpenExcelWorkbook(tclist_filename, XLS: true);
                        MainForm.SystemLogAddLine("TestCase file is XML");
                    }
                    else
                    {
                        wb_tc = ExcelAction.OpenExcelWorkbook(tclist_filename);
                        MainForm.SystemLogAddLine("TestCase file is XLS(X)");
                    }
                }
                else
                {
                    wb_tc = ExcelAction.OpenExcelWorkbook(tclist_filename);
                    MainForm.SystemLogAddLine("TC template file is XLS(X)");
                }

                if (wb_tc == null)
                {
                    return ExcelStatus.ERR_OpenTestCaseExcel_OpenExcelWorkbook;
                }

                Worksheet ws_tclist;
                if (IsTemplate == false)
                {
                    ws_tclist = ExcelAction.Find_Worksheet(wb_tc, TestCase.TestCaseSheetName);
                }
                else
                {
                    ws_tclist = ExcelAction.Find_Worksheet(wb_tc, TestCase.TCTemplateSheetName);
                }

                if (ws_tclist == null)
                {
                    if (IsTemplate == false)
                    {
                        MainForm.SystemLogAddLine("TestCase file does not have specified sheetname");
                    }
                    else
                    {
                        MainForm.SystemLogAddLine("TC template does not have specified sheetname");
                    }
                    return ExcelStatus.ERR_OpenTestCaseExcel_Find_Worksheet;
                }
                else
                {
                    if (IsTemplate == false)
                    {
                        workbook_testcase = wb_tc;
                        ws_testcase = ws_tclist;
                    }
                    else
                    {
                        workbook_tc_template = wb_tc;
                        ws_tc_template = ws_tclist;
                    }
                    return ExcelStatus.OK;
                }
            }
            catch
            {
                MainForm.SystemLogAddLine("OpenTestCaseExcel has exception");
                return ExcelStatus.EX_OpenTestCaseWorksheet;
            }

            // Not needed because never reaching here
            //return ExcelStatus.ERR_NOT_DEFINED;
        }

        static public ExcelStatus CloseTestCaseExcel(bool IsTemplate = false)
        {
            try
            {
                if (IsTemplate == false)
                {
                    if (workbook_testcase == null)
                    {
                        return ExcelStatus.ERR_CloseTestCaseExcel_wb_null;
                    }
                    ExcelAction.CloseExcelWorkbook(workbook_testcase, SaveChanges: false);
                    ws_testcase = null;
                    workbook_testcase = null;
                }
                else
                {
                    if (workbook_tc_template == null)
                    {
                        return ExcelStatus.ERR_CloseTestCaseExcel_wb_null;
                    }
                    ExcelAction.CloseExcelWorkbook(workbook_tc_template, SaveChanges: false);
                    ws_tc_template = null;
                    workbook_tc_template = null;
                }
                return ExcelStatus.OK;
            }
            catch
            {
                return ExcelStatus.EX_CloseTestCaseWorksheet;
            }
        }

        static public ExcelStatus SaveChangesAndCloseTestCaseExcel(String dest_filename, bool IsTemplate = false)
        {
            try
            {
                if (IsTemplate == false)
                {
                    if (workbook_testcase == null)
                    {
                        return ExcelStatus.ERR_SaveChangesAndCloseTestCaseExcel_wb_null;
                    }
                    ExcelAction.CloseExcelWorkbook(workbook_testcase, SaveChanges: true, AsFilename: dest_filename);
                    ws_testcase = null;
                    workbook_testcase = null;
                }
                else
                {
                    if (workbook_tc_template == null)
                    {
                        return ExcelStatus.ERR_SaveChangesAndCloseTestCaseExcel_wb_null;
                    }
                    ExcelAction.CloseExcelWorkbook(workbook_tc_template, SaveChanges: true, AsFilename: dest_filename);
                    ws_tc_template = null;
                    workbook_tc_template = null;
                }
                return ExcelStatus.OK;
            }
            catch
            {
                return ExcelStatus.EX_SaveChangesAndCloseTestCaseExcel;
            }
            // Not needed because never reaching here
            //return ExcelStatus.ERR_NOT_DEFINED;
        }

        // Copy Value2 of single cell or a range of cells

        static private void CopyValue2(Worksheet src, Worksheet dst, int ul_row, int ul_col)
        {
            Range Src = src.Cells[ul_row, ul_col];
            Range Dst = dst.Cells[ul_row, ul_col];
            Dst.Value2 = Src.Value2;
        }

        static private void CopyValue2_different_cell_location(Worksheet src, int src_row, int src_col,
                                                            Worksheet dst, int dst_row, int dst_col)
        {
            Range Src = src.Cells[src_row, src_col];
            Range Dst = dst.Cells[dst_row, dst_col];
            Dst.Value2 = Src.Value2;
        }

        static private void CopyValue2_different_range_location(Worksheet src, int src_ul_row, int src_ul_col, int src_br_row, int src_br_col,
                                                                Worksheet dst, int dst_ul_row, int dst_ul_col)
        {
            int dst_br_row = dst_ul_row + (src_br_row - src_ul_row),
                dst_br_col = dst_ul_col + (src_br_col - src_ul_col);
            Range Src = src.Range[src.Cells[src_ul_row, src_ul_col], src.Cells[src_br_row, src_br_col]];
            Range Dst = dst.Range[dst.Cells[dst_ul_row, dst_ul_col], dst.Cells[dst_br_row, dst_br_col]];
            Dst.Value2 = Src.Value2;
        }

        static private void CopyValue2(Worksheet src, Worksheet dst, int ul_row, int ul_col, int br_row, int br_col)
        {
            Range Src = src.Range[src.Cells[ul_row, ul_col], src.Cells[br_row, br_col]];
            Range Dst = dst.Range[dst.Cells[ul_row, ul_col], dst.Cells[br_row, br_col]];
            Dst.Value2 = Src.Value2;
        }

        static private void CopyPaste(Worksheet src, Worksheet dst, int ul_row, int ul_col)
        {
            CopyPaste(src, dst, ul_row, ul_col, ul_row, ul_col);
        }

        static private void CopyPaste(Worksheet src, Worksheet dst, int ul_row, int ul_col, int br_row, int br_col)
        {
            Range Src = src.Range[src.Cells[ul_row, ul_col], src.Cells[br_row, br_col]];
            Range Dst = dst.Range[dst.Cells[ul_row, ul_col], dst.Cells[br_row, br_col]];
            Src.Copy();
            Dst.PasteSpecial(Paste: XlPasteType.xlPasteAll);
        }

        static public void CopyPasteRows(Worksheet src, Worksheet dst, int start_row, int end_row)
        {
            String row_str = start_row.ToString() + ":" + end_row.ToString();
            Range Src = src.Rows[row_str];
            Range Dst = dst.Rows[row_str];
            Src.Copy();
            Dst.PasteSpecial(Paste: XlPasteType.xlPasteAll);
        }
        static private void CopyPasteFormat(Worksheet src, Worksheet dst, int ul_row, int ul_col)
        {
            CopyPasteFormat(src, dst, ul_row, ul_col, ul_row, ul_col);
        }

        static private void CopyPasteFormat(Worksheet src, Worksheet dst, int ul_row, int ul_col, int br_row, int br_col)
        {
            Range Src = src.Range[src.Cells[ul_row, ul_col], src.Cells[br_row, br_col]];
            Range Dst = dst.Range[dst.Cells[ul_row, ul_col], dst.Cells[br_row, br_col]];
            Src.Copy();
            Dst.PasteSpecial(Paste: XlPasteType.xlPasteFormats);
        }

        static public void SetFontColorToWhite(Worksheet worksheet, int ul_row, int ul_col, int br_row, int br_col)
        {
            Range range = worksheet.Range[worksheet.Cells[ul_row, ul_col], worksheet.Cells[br_row, br_col]];
            range.Characters.Font.Color = System.Drawing.Color.White; ;
        }

        /*
        // Copy value2 of Test-Case Excel (tc_data) to Test-Case-Template Excel.
        // Result: Test-Case Excel data shown in the format of Test-Case-Template
        static public bool CopyTestCaseIntoTemplate()
        {
            Worksheet tc_data = ws_testcase,
                      tc_template = ws_tc_template;

            // Protection
            if (tc_data == null) { return false; }
            if (tc_template == null) { return false; }

            Worksheet ws_Src = tc_data, ws_Dst = tc_template;
            Range Src = GetTestCaseAllRange();
            Range Dst = GetTestCaseAllRange(IsTemplate: true);
            int Src_last_row = Src.Row, Src_last_col = Src.Column;
            int Dst_last_row = Dst.Row, Dst_last_col = Dst.Column;

            // workaround for temp

            // Make template (destination) row count == TestCase (source) row count
            if (Src_last_row > Dst_last_row)
            {
                // Insert row into template file
                int rows_to_insert = Src_last_row - Dst_last_row;
                do
                {
                    Insert_Row(ws_Dst, TestCase.DataBeginRow + 1);
                }
                while (--rows_to_insert > 0);
            }
            else if (Src_last_row < Dst_last_row)
            {
                // Delete row from template file
                int rows_to_delete = Dst_last_row - Src_last_row;
                do
                {
                    Delete_Row(ws_Dst, TestCase.DataBeginRow);
                }
                while (--rows_to_delete > 0);
            }

            // Copy [3,1] from tc to template
            CopyValue2(ws_Src, ws_Dst, 3, 1);

            // Copy row 4 (Column Name) from tc to template
            CopyValue2(ws_Src, ws_Dst, TestCase.NameDefinitionRow, 1, TestCase.NameDefinitionRow, Src_last_col);

            // Copy [Src_last_row,1] from tc to template
            CopyValue2(ws_Src, ws_Dst, Src_last_row, 1);

            // Copy the rest of data
            CopyValue2(ws_Src, ws_Dst, TestCase.DataBeginRow, 1, Src_last_row - 1, Src_last_col);

            return true;
        }
        */

        // This version doesn't assume that columns item/sequence are both the same.
        // 1. adjust rows of tc_template to be the same as test-case excel (by inserting/deleting rows)
        // 2. don't copy column name --> keep them instead
        // 3. copy cell value (of each data row) into correpsonding column of tc_template.
        //     for example copy "Key" (assumed at column 1) to another "Key" (assumed not at column 1)
        static public Boolean CopyTestCaseIntoTemplate_v2()
        {
            Worksheet tc_data = ws_testcase,
                      tc_template = ws_tc_template;

            // Protection
            if (tc_data == null) { return false; }
            if (tc_template == null) { return false; }

            Worksheet ws_Src = tc_data, ws_Dst = tc_template;
            Range Src = GetWorksheetAllRange(ws_Src);
            Range Dst = GetWorksheetAllRange(ws_Dst);
            int Src_last_row = Get_Range_RowNumber(Src), Src_last_col = Get_Range_ColumnNumber(Src);
            int Dst_last_row = Get_Range_RowNumber(Dst), Dst_last_col = Get_Range_ColumnNumber(Dst);

            // Neither insert nor remove rows now. template file is responsible for setting up desired format
            //// Make template (destination) row count == TestCase (source) row count
            //if (Src_last_row > Dst_last_row)
            //{
            //    // Insert row into template file
            //    int rows_to_insert = Src_last_row - Dst_last_row;
            //    do
            //    {
            //        Insert_Row(ws_Dst, TestCase.DataBeginRow + 1);
            //    }
            //    while (--rows_to_insert > 0);
            //}
            //else if (Src_last_row < Dst_last_row)
            //{
            //    // Delete row from template file
            //    int rows_to_delete = Dst_last_row - Src_last_row;
            //    do
            //    {
            //        Delete_Row(ws_Dst, TestCase.DataBeginRow);
            //    }
            //    while (--rows_to_delete > 0);
            //}

            // Copy [3,1] from tc to template
            //CopyValue2(ws_Src, ws_Dst, 3, 1);

            // Copy row 4 (Column Name) from tc to template
            //CopyValue2(ws_Src, ws_Dst, TestCase.NameDefinitionRow, 1, TestCase.NameDefinitionRow, Src_last_col);

            // Copy [Src_last_row,1] from tc to template
            //CopyValue2(ws_Src, ws_Dst, Src_last_row, 1);

            // Copy the rest of data
            //CopyValue2(ws_Src, ws_Dst, TestCase.DataBeginRow, 1, Src_last_row - 1, Src_last_col);

            // use LUT of column index for mapping the same column_name of SRC/DST
            Dictionary<string, int> src_col_name_list = TestCase.TestCaseColumnIndexLUT();
            Dictionary<string, int> dst_col_name_list = TestCase.TCTemplateColumnIndexLUT();
            int source_row_begin = TestCase.TC_DataBeginRow, 
                destination_row_begin = TestCase.Template_DataBeginRow, 
                source_row_end = Src_last_row;

            // reduce row_end by 1 when there isn't valid key value at last row.
            String check_key, check_summary;
            check_key = GetCellTrimmedString(ws_Src, source_row_end, src_col_name_list[TestCase.col_Key]);
            check_summary = GetCellTrimmedString(ws_Src, source_row_end, src_col_name_list[TestCase.col_Summary]);
            if (TestCase.CheckValidTC_By_Key_Summary(check_key, check_summary) == false)
            {
                source_row_end--;
            }

            foreach (string col_name in dst_col_name_list.Keys)
            {
                // check to prevent outputing unavailable columns in source to destination
                if (src_col_name_list.ContainsKey(col_name))
                {
                    int dst_col = dst_col_name_list[col_name], src_col = src_col_name_list[col_name];
                    CopyValue2_different_range_location(ws_Src, source_row_begin, src_col, source_row_end, src_col, ws_Dst, destination_row_begin, dst_col);
                }
            }

            return true;
        }

        //static public ExcelStatus OpenKeywordLogTemplateExcel(String template_filename)
        //{
        //    try
        //    {
        //        Workbook wb_keywordlog;

        //        // Open excel (read-only & corrupt-load)
        //        wb_keywordlog = ExcelAction.OpenExcelWorkbook(template_filename);

        //        if (wb_keywordlog == null)
        //        {
        //            return ExcelStatus.ERR_OpenKeywordLogTemplateExcel_OpenExcelWorkbook;
        //        }

        //        // Check both worksheet
        //        Worksheet ws_kw_list = ExcelAction.Find_Worksheet(wb_keywordlog, KeyWordListReport.WS_KeyWord_List);
        //        if (ws_kw_list == null)
        //        {
        //            return ExcelStatus.ERR_OpenKeywordLogTemplateExcel_Find_Keyword_Worksheet;
        //        }
        //        Worksheet ws_not_kw_file = ExcelAction.Find_Worksheet(wb_keywordlog, KeyWordListReport.WS_NotKeyWord_File);
        //        if (ws_not_kw_file == null)
        //        {
        //            return ExcelStatus.ERR_OpenKeywordLogTemplateExcel_Find_NonKeyword_Worksheet;
        //        }

        //        workbook_keywordlog_template = wb_keywordlog;
        //        ws_keyword_list = ws_kw_list;
        //        ws_not_keyword_file = ws_not_kw_file;

        //        return ExcelStatus.OK;
        //    }
        //    catch
        //    {
        //        return ExcelStatus.EX_OpenKeywordLogTemplateWorksheet;
        //    }

        //    // Not needed because never reaching here
        //    //return ExcelStatus.ERR_NOT_DEFINED;
        //}

        //static public ExcelStatus CloseKeywordLogTemplateExcel()
        //{
        //    try
        //    {
        //        if (workbook_keywordlog_template == null)
        //        {
        //            return ExcelStatus.ERR_CloseKeywordLogTemplateExcel_wb_null;
        //        }
        //        ExcelAction.CloseExcelWorkbook(workbook_keywordlog_template, SaveChanges: false);
        //        ws_keyword_list = ws_not_keyword_file = null;
        //        workbook_keywordlog_template = null;
        //        return ExcelStatus.OK;
        //    }
        //    catch
        //    {
        //        ws_keyword_list = ws_not_keyword_file = null;
        //        workbook_keywordlog_template = null;
        //        return ExcelStatus.EX_CloseKeywordLogTemplateExcel;
        //    }
        //}

        //static public ExcelStatus SaveChangesAndCloseKeywordLogTemplateExcel(String dest_filename)
        //{
        //    try
        //    {
        //        if (workbook_keywordlog_template == null)
        //        {
        //            return ExcelStatus.ERR_SaveChangesAndCloseKeywordLogTemplateExcel_wb_null;
        //        }
        //        ExcelAction.CloseExcelWorkbook(workbook_keywordlog_template, SaveChanges: true, AsFilename: dest_filename);
        //        ws_keyword_list = ws_not_keyword_file = null;
        //        workbook_keywordlog_template = null;
        //        return ExcelStatus.OK;
        //    }
        //    catch
        //    {
        //        ws_keyword_list = ws_not_keyword_file = null;
        //        workbook_keywordlog_template = null;
        //        return ExcelStatus.EX_SaveChangesAndCloseKeywordLogTemplateExcel;
        //    }
        //}

        static public ExcelStatus CreateNewKeywordListExcel()
        {
            int original_SheetsInNewWorkbook = excel_app.SheetsInNewWorkbook;

            excel_app.SheetsInNewWorkbook = 2;

            Workbook wb = excel_app.Workbooks.Add(Missing.Value);
            workbook_new_keyword_list = wb;

            ws_keyword_list = workbook_new_keyword_list.Sheets.Item[1];
            ws_keyword_list.Name = KeyWordListReport.WS_KeyWord_List;

            ws_not_keyword_file = workbook_new_keyword_list.Sheets.Item[2];
            ws_not_keyword_file.Name = KeyWordListReport.WS_NotKeyWord_File;

            excel_app.SheetsInNewWorkbook = original_SheetsInNewWorkbook;
            return ExcelStatus.OK;
        }

        static public ExcelStatus CloseNewKeywordListExcel()
        {
            try
            {
                if (workbook_new_keyword_list == null)
                {
                    return ExcelStatus.ERR_CloseKeywordLogTemplateExcel_wb_null;
                }
                ExcelAction.CloseExcelWorkbook(workbook_new_keyword_list, SaveChanges: false);
                ws_keyword_list = ws_not_keyword_file = null;
                workbook_new_keyword_list = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_keyword_list = ws_not_keyword_file = null;
                workbook_new_keyword_list = null;
                return ExcelStatus.EX_CloseKeywordLogTemplateExcel;
            }
        }

        static public ExcelStatus SaveChangesAndCloseNewKeywordListExcel(String dest_filename)
        {
            try
            {
                if (workbook_new_keyword_list == null)
                {
                    return ExcelStatus.ERR_SaveChangesAndCloseKeywordLogTemplateExcel_wb_null;
                }
                ExcelAction.CloseExcelWorkbook(workbook_new_keyword_list, SaveChanges: true, AsFilename: dest_filename);
                ws_keyword_list = ws_not_keyword_file = null;
                workbook_new_keyword_list = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_keyword_list = ws_not_keyword_file = null;
                workbook_new_keyword_list = null;
                return ExcelStatus.EX_SaveChangesAndCloseKeywordLogTemplateExcel;
            }
        }

        static public String default_table_font_name = "Mabry Pro";
        static public int default_table_font_size = 12;
        static public Color default_table_font_color = Color.Black;
        static public FontStyle default_table_font_style = FontStyle.Regular;

        static public void WriteTableObjectToExcel(Worksheet worksheet, List<List<Object>> table_object,
                            int start_row = 1, int start_col = 1, Boolean with_title = true,
                            List<int> left_alignment_col = null, List<int> center_alignment_col = null,
                            List<int> auto_fit_col = null)
        {
            int row_pos = start_row;
            int col_pos = start_col;
            int row_end = start_row;
            int col_end = start_col;
            int content_start_row = start_row;

            // 1. Fill worksheet with objects
            foreach (List<Object> row_obj_list in table_object)
            {
                foreach (Object obj in row_obj_list)
                {
                    SetCellValue(worksheet, row_pos, col_pos++, obj);
                }
                // update new right border of table
                if ((col_pos - 1) > col_end)
                {
                    col_end = col_pos - 1;
                }
                row_pos++;
                col_pos = start_col;
            }
            row_end = row_pos - 1;

            // 2. formating all table cells with font & border / auto-fit columns
            Range table_range = worksheet.Range[worksheet.Cells[start_row, start_col], worksheet.Cells[row_end, col_end]];
            table_range.NumberFormat = "@";
            using (System.Drawing.Font fontTester = new System.Drawing.Font(StyleString.default_font, StyleString.default_size,
                                                StyleString.default_fontstyle, GraphicsUnit.Pixel))
            {
                if (fontTester.Name == StyleString.default_font)
                {
                    // Font exists
                    table_range.Characters.Font.Name = StyleString.default_font;
                    table_range.Characters.Font.Size = StyleString.default_size;
                    table_range.Characters.Font.Color = StyleString.default_color;
                    table_range.Characters.Font.FontStyle = StyleString.default_fontstyle;
                }
                else
                {
                    // Font doesn't exist ==> use internal default
                    table_range.Characters.Font.Name = default_table_font_name;
                    table_range.Characters.Font.Size = default_table_font_size;
                    table_range.Characters.Font.Color = default_table_font_color;
                    table_range.Characters.Font.FontStyle = default_table_font_style;
                }
            }
            table_range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            table_range.Borders.Weight = Excel.XlBorderWeight.xlThin;
            table_range.NumberFormat = "@";

            // 3. format cell BG color if with_title is true
            if (with_title)
            {
                Range title_range = worksheet.Range[worksheet.Cells[start_row, start_col], worksheet.Cells[start_row, col_end]];
                title_range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                title_range.Interior.Color = Color.LightGray;
                content_start_row++; // adjust content_start_row to exclude title row in the following operation
            }

            // 4.auto-fit columns
            if (auto_fit_col != null)
            {
                foreach (int col in auto_fit_col)
                {
                    AutoFit_Column(worksheet, col);
                }
            }

            // 5. left-alignment specific column 
            if (center_alignment_col != null)
            {
                foreach (int col in left_alignment_col)
                {
                    Range col_range = worksheet.Range[worksheet.Cells[content_start_row, col], worksheet.Cells[row_end, col]];
                    col_range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }
            }

            // 6. center-alignment specific column 
            if (center_alignment_col != null)
            {
                foreach (int col in center_alignment_col)
                {
                    Range col_range = worksheet.Range[worksheet.Cells[content_start_row, col], worksheet.Cells[row_end, col]];
                    col_range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                }
            }
        }

        static public void WriteTableToKeywordList(List<List<Object>> table_object)
        {
            List<int> left_alignment_col = new List<int>();
            List<int> center_alignment_col = new List<int>();
            List<int> auto_fit_col = new List<int>();
            int start_row = KeyWordListReport.keyword_list_title_row,
                start_col = KeyWordListReport.keyword_list_title_col_start;
            Boolean with_title = true;

            left_alignment_col.Add(1);
            left_alignment_col.Add(2);
            left_alignment_col.Add(3);
            left_alignment_col.Add(4);
            center_alignment_col.Add(5);
            auto_fit_col.Add(1);
            auto_fit_col.Add(3);
            auto_fit_col.Add(4);
            auto_fit_col.Add(5);
            WriteTableObjectToExcel(ws_keyword_list, table_object, start_row, start_col, with_title,
                                    left_alignment_col, center_alignment_col, auto_fit_col);
        }

        static public void WriteTableToNotKeywordFile(List<List<Object>> table_object)
        {
            List<int> left_alignment_col = new List<int>();
            List<int> center_alignment_col = new List<int>();
            List<int> auto_fit_col = new List<int>();
            int start_row = KeyWordListReport.not_keyword_file_title_row,
                start_col = KeyWordListReport.not_keyword_file_title_col_start;
            Boolean with_title = true;

            left_alignment_col.Add(1);
            left_alignment_col.Add(2);
            center_alignment_col.Add(3);
            center_alignment_col.Add(4);
            center_alignment_col.Add(5);
            center_alignment_col.Add(6);
            auto_fit_col.Add(2);
            auto_fit_col.Add(3);
            auto_fit_col.Add(4);
            auto_fit_col.Add(5);
            auto_fit_col.Add(6);
            WriteTableObjectToExcel(ws_not_keyword_file, table_object, start_row, start_col, with_title,
                                    left_alignment_col, center_alignment_col, auto_fit_col);
        }

        static public void WriteTestReportCreationLog(List<List<Object>> table_object)
        {
            // TC-Test Group, TC-Summary, source_path, file_name, not_found, not_copied, dest_path
            // 

            List<int> left_alignment_col = new List<int>();
            List<int> center_alignment_col = new List<int>();
            List<int> auto_fit_col = new List<int>();
            int start_row = KeyWordListReport.not_keyword_file_title_row,
                start_col = KeyWordListReport.not_keyword_file_title_col_start;
            Boolean with_title = true;

            left_alignment_col.Add(1);
            left_alignment_col.Add(2);
            center_alignment_col.Add(3);
            center_alignment_col.Add(4);
            center_alignment_col.Add(5);
            center_alignment_col.Add(6);
            auto_fit_col.Add(2);
            auto_fit_col.Add(3);
            auto_fit_col.Add(4);
            auto_fit_col.Add(5);
            auto_fit_col.Add(6);
            WriteTableObjectToExcel(ws_not_keyword_file, table_object, start_row, start_col, with_title,
                                    left_alignment_col, center_alignment_col, auto_fit_col);
        }

        static public void ClearContent(Worksheet worksheet, int start_row, int start_col, int end_row, int end_col)
        {
            Range table_range = worksheet.Range[worksheet.Cells[start_row, start_col], worksheet.Cells[end_row, end_col]];
            table_range.Value = " ";
            table_range.Interior.ColorIndex = 0;
            table_range.Characters.Font.Color = StyleString.default_color;
        }

        static public void SetBGColorRow(Worksheet worksheet, int start_row, int end_row, Color bg_color)
        {
            int countColumns = worksheet.UsedRange.Columns.Count;
            Range table_range = worksheet.Range[worksheet.Cells[start_row, 1], worksheet.Cells[end_row, countColumns]];
            table_range.Interior.Color = bg_color;
        }

        static public void Merge(Worksheet worksheet, int upper_row, int left_column, int bottom_row, int right_column)
        {
            Range merge_range = worksheet.Range[worksheet.Cells[upper_row, left_column], worksheet.Cells[bottom_row, right_column]];
            try
            {
                merge_range.Merge(true);
            }
            catch (Exception ex)
            {
                Error_Log.Add("Merge exception of (" + ex.ToString() + ") at ared of: " + upper_row.ToString() + "," + left_column.ToString() + "," + bottom_row.ToString() + "," + right_column.ToString()); // here is for setting break-point
                // already merged?
            }
        }

        static public void UnMerge(Worksheet worksheet, int upper_row, int left_column, int bottom_row, int right_column)
        {
            Range merge_range = worksheet.Range[worksheet.Cells[upper_row, left_column], worksheet.Cells[bottom_row, right_column]];
            merge_range.UnMerge();
        }

        // Return the column number for this column name.
        static public int ColumnNameToNumber(String col_name)
        {
            int result = 0;

            // Process each letter.
            for (int i = 0; i < col_name.Length; i++)
            {
                result *= 26;
                char letter = col_name[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }
            return result;
        }

        static public int ColumnNameToNumber(Char col_char)
        {
            int result = 0;

            if (col_char < 'A')
                col_char = 'A';
            if (col_char > 'Z')
                col_char = 'Z';

            result = (int)col_char - (int)'A' + 1;
            return result;
        }

        //
        // Copy bug-list into TC template
        // NOTE: bug-list sheetname has been renamed before entering this function call
        //
        static public Boolean CopyBugListSheetIntoTestCaseTemplateWorkbook()
        {
            Boolean b_ret = false;
            Worksheet template_last_sheet = workbook_tc_template.Worksheets[workbook_tc_template.Worksheets.Count];

            int index_copy_bug_list = 0, index_copy_and_extend_bug_list = 0;                        // 0 means not copied
            Worksheet worksheet_copy_bug_list = null, worksheet_copy_and_extend_bug_list = null;
            int index_added_after_worksheet = ws_tc_template.Index;

            try
            {
                if (ReportGenerator.copy_bug_list) // original bug-list are copied without modificaiton
                {
                    ws_issuelist.Copy(After: workbook_tc_template.Worksheets[index_added_after_worksheet]);
                    index_copy_bug_list = ++index_added_after_worksheet;
                    worksheet_copy_bug_list = workbook_tc_template.Worksheets[index_copy_bug_list];
                }

                if (ReportGenerator.copy_and_extend_bug_list)   // bug-list are copied and linked issue (Test-case here) are extended from "Key" into "Summary" of testcase
                {
                    ws_issuelist.Copy(After: workbook_tc_template.Worksheets[index_added_after_worksheet]);
                    index_copy_and_extend_bug_list = ++index_added_after_worksheet;
                    worksheet_copy_and_extend_bug_list = workbook_tc_template.Worksheets[index_copy_and_extend_bug_list];
                    ReportGenerator.ProcessBugListToExtendTestCase(worksheet_copy_and_extend_bug_list);
                }
                b_ret = true;
            }
            catch (Exception ex)
            {
            }

            return b_ret;
        }

        static public Boolean DuplicateReportListSheet(Worksheet ws_reportlist)
        {
            Boolean b_ret = false;
            try
            {
                ws_reportlist.Copy(After: ws_reportlist);
                b_ret = true;
            }
            catch (Exception ex)
            {
            }
            return b_ret;
        }

        // Assumption: (as default value)
        //  1. Report to be copied is sheet 1
        //  2. copied to sheet 2, sheetname is month/date/hour/minute
        static public Boolean CopyReportSheetAsHistory(Workbook workbook_report, String sheetname_append_string, int ReportIndex = 1, int DestinationIndex = 2)
        {
            Boolean b_ret = false;

            int current_worksheet_count = workbook_report.Worksheets.Count;
            DestinationIndex--;   // decide inserting after which sheet. if minus-one ==0, it means inserting at the beginning
            if ((ReportIndex > 0) && (ReportIndex <= current_worksheet_count) && (DestinationIndex >= 0) && (DestinationIndex <= (current_worksheet_count)))
            {
                Worksheet source = workbook_report.Worksheets[ReportIndex];
                if (DestinationIndex == 0)      // insert at the beginning to become 1st worksheet
                {
                    source.Copy(Before: workbook_report.Worksheets[1]);
                }
                else
                {
                    source.Copy(After: workbook_report.Worksheets[DestinationIndex]);
                }
                Worksheet copied = workbook_report.Worksheets[++DestinationIndex];
                copied.Name = source.Name + sheetname_append_string;
                source.Select();
                b_ret = true;
            }

            return b_ret;
        }
        static public ExcelData InitTCExcelData(bool IsTemplate = false)
        {
            ExcelData excel_data = new ExcelData();
            Worksheet ws_tclist;
            int nameDefinitionRow, dataBeginRow;
            if (IsTemplate)
            {
                ws_tclist = ws_tc_template;
                nameDefinitionRow = TestCase.Template_NameDefinitionRow;
                dataBeginRow = TestCase.Template_DataBeginRow;
            }
            else
            {
                ws_tclist = ws_testcase;
                nameDefinitionRow = TestCase.TC_NameDefinitionRow;
                dataBeginRow = TestCase.TC_DataBeginRow;
            }
            excel_data.InitFromExcel(ws_tclist, nameDefinitionRow, dataBeginRow);
            return excel_data;
        }

        //static public Workbook OpenCSV(String filename)
        //{
        //    Workbook ret_workbook = null;
        //    // Protection-check: file must exist.
        //    if (!Storage.FileExists(filename))
        //    {
        //        return ret_workbook;
        //    }

        //    try
        //    {
        //        ret_workbook = excel_app.Workbooks.Open(
        //            filename,               // Filename
        //            Type.Missing, Type.Missing,
        //            XlFileFormat.xlCSV,   // Format
        //            Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
        //            ",",          // Delimiter
        //            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //    }
        //    catch (Exception ex)
        //    {
        //        // 
        //    }

        //    return ret_workbook;
        //}

        //static public void CloseCSV_SaveAsExcel(Workbook workbook, bool SaveChanges = false, String AsFilename = "")
        //{
        //    excel_app.DisplayAlerts = false;
        //    if (SaveChanges)
        //    {
        //        if (String.IsNullOrWhiteSpace(AsFilename) == false)
        //        if (AsFilename != "")
        //        {
        //            workbook.SaveAs(AsFilename, Type.Missing,
        //                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
        //             Type.Missing, Type.Missing, Type.Missing, true, Type.Missing);
        //            workbook.Close(SaveChanges: false);
        //        }
        //        else
        //        {
        //            workbook.Save();
        //            workbook.Close(SaveChanges: false);
        //        }
        //    }
        //    else
        //    {
        //        workbook.Close(SaveChanges: false);
        //    }
        //}
    }
}
