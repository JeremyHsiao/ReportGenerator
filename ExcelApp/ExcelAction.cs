using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelReportApplication
{
    static class ExcelAction
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

        static public bool ExcelVisible = true;

        static public void OpenExcelApp()
        {
            if (excel_app != null) return;
            excel_app = new Excel.Application();
            excel_app.Visible = ExcelVisible;
            excel_app.Caption = "DQA Report Generator";
            excel_app.DisplayAlerts = false;
        }

        static public Workbook OpenExcelWorkbook(String filename, bool ReadOnly = true, bool XLS = false, bool UpdateLinks = false)
        {
            Workbook ret_workbook;
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
            return ret_workbook;
        }

        static public void CloseExcelWorkbook(Workbook workbook, bool SaveChanges = false, String AsFilename = "")
        {
            excel_app.DisplayAlerts = false;
            if (SaveChanges)
            {
                if (AsFilename != "")
                {
                    workbook.Close(SaveChanges: true, Filename: AsFilename);
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
            foreach (Excel.Worksheet displayWorksheet in curExcel.Worksheets)
            {
                ret_sheetname.Add(displayWorksheet.Name);
            }
            return ret_sheetname;
        }

        static public bool WorksheetExist(Excel.Application curExcel, string sheet_name)
        {
            foreach (Excel.Worksheet displayWorksheet in curExcel.Worksheets)
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
            foreach (Excel.Worksheet displayWorksheet in wb.Worksheets)
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
            return ws.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
        }

        static public Range GetWorksheetPrintableRange(Worksheet ws)
        {
            String PrintArea = ws.PageSetup.PrintArea;
            Range rngPrintable = ws.Range[PrintArea];
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

        static public void Hide_Row(Worksheet ws, int row, int count = 1)
        {
            Range hiddenRange = ws.Range[ws.Cells[row, 1], ws.Cells[row + count - 1, GetWorksheetAllRange(ws).Column]];
            //     var hiddenRange = yourWorksheet.Range[yourWorksheet.Cells[firstRowToHide, firstColToHide], yourWorksheet.Cells[lastRowToHide, lastColToHide]];
            hiddenRange.EntireRow.Hidden = true;
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

        static public Dictionary<string, int> CreateTableColumnIndex(Worksheet ws, int naming_row)
        {
            Dictionary<string, int> col_name_list = new Dictionary<string, int>();

            int col_end = GetWorksheetAllRange(ws).Column;
            for (int col_index = 1; col_index <= col_end; col_index++)
            {
                String cell_value2 = GetCellTrimmedString(ws, naming_row, col_index);
                if (cell_value2 == "") { continue; }
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
            ERR_CloseIssueListExcel_wb_null,
            ERR_CloseTestCaseExcel_wb_null,
            ERR_SaveChangesAndCloseIssueListExcel_wb_null,
            ERR_SaveChangesAndCloseTestCaseExcel_wb_null,
            ERR_NOT_DEFINED,
            EX_OpenIssueListWorksheet,
            EX_CloseIssueListExcel,
            EX_SaveChangesAndCloseIssueListExcel,
            EX_OpenTestCaseWorksheet,
            EX_CloseTestCaseWorksheet,
            EX_SaveChangesAndCloseTestCaseExcel,
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

        static public void IssueList_WriteStyleString(int row, int col, List<StyleString> sytle_string_list)
        {
            StyleString.WriteStyleString(ws_issuelist, row, col, sytle_string_list);
        }

        static public Dictionary<string, int> CreateIssueListColumnIndex()
        {
            return CreateTableColumnIndex(ws_issuelist, Issue.NameDefinitionRow);
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

        static public String GetTestCaseCellTrimmedString(int row, int col, bool IsTemplate = false)
        {
            Object cell_value2 = GetTestCaseCell(row, col, IsTemplate: IsTemplate);
            if (cell_value2 == null) { return ""; }
            return cell_value2.ToString();
        }

        static public Object GetCellValue(Worksheet ws, int row, int col)
        {
            return ws.Cells[row, col].Value2;
        }

        static public void SetCellValue(Worksheet ws, int row, int col, Object value)
        {
            ws.Cells[row, col].Value2 = value;
        }

        static public String GetCellTrimmedString(Worksheet ws, int row, int col)
        {
            Object cell_value2 = GetCellValue(ws, row, col);
            if (cell_value2 == null) { return ""; }
            return cell_value2.ToString();
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

        static public void TestCase_WriteStyleString(int row, int col, List<StyleString> sytle_string_list, bool IsTemplate = false)
        {
            Worksheet ws = ((IsTemplate) ? ws_tc_template : ws_testcase);
            StyleString.WriteStyleString(ws, row, col, sytle_string_list);
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

        static public Dictionary<string, int> CreateTestCaseColumnIndex(bool IsTemplate = false)
        {
            return CreateTableColumnIndex(((IsTemplate)?ws_tc_template:ws_testcase), TestCase.NameDefinitionRow);
        }

        // Excel Open/Close/Save for Issue List Excel

        static public ExcelStatus OpenIssueListExcel(String buglist_filename)
        {
            try
            {
                Workbook wb_issuelist;

                // Open excel (read-only & corrupt-load)
                wb_issuelist = ExcelAction.OpenExcelWorkbook(buglist_filename, XLS: true);

                if (wb_issuelist == null)
                {
                    return ExcelStatus.ERR_OpenIssueListExcel_OpenExcelWorkbook;
                }

                Worksheet ws_buglist = ExcelAction.Find_Worksheet(wb_issuelist, Issue.SheetName);
                if (ws_buglist == null)
                {
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
        }

        // Excel Open/Close/Save for Test Case Excel

        static public ExcelStatus OpenTestCaseExcel(String tclist_filename, bool IsTemplate = false)
        {
            try
            {
                Workbook wb_tc;
                if (IsTemplate == false)
                {
                    // Open excel (read-only & corrupt-load)
                    wb_tc = ExcelAction.OpenExcelWorkbook(tclist_filename, XLS: true);
                }
                else
                {
                    wb_tc = ExcelAction.OpenExcelWorkbook(tclist_filename);
                }

                if (wb_tc == null)
                {
                    return ExcelStatus.ERR_OpenTestCaseExcel_OpenExcelWorkbook;
                }

                Worksheet ws_tclist = ExcelAction.Find_Worksheet(wb_tc, TestCase.SheetName);
                if (ws_tclist == null)
                {
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
        }

        // Copy Value2 of single cell or a range of cells

        static private void CopyValue2(Worksheet src, Worksheet dst, int ul_row, int ul_col)
        {
            Range Src = src.Cells[ul_row, ul_col];
            Range Dst = dst.Cells[ul_row, ul_col];
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

        // Copy value2 of Test-Case Excel (tc_data) to Test-Case-Template Excel.
        // Result: Test-Case Excel data shown in the format of Test-Case-Template
        static public bool CopyTestCaseIntoTemplate()
        {
            Worksheet tc_data       = ws_testcase,
                      tc_template = ws_tc_template;

            // Protection
            if (tc_data == null) { return false; }
            if (tc_template == null) { return false; }

            Worksheet ws_Src = tc_data, ws_Dst = tc_template;
            Range Src = GetWorksheetAllRange(ws_Src);
            Range Dst = GetWorksheetAllRange(ws_Dst);
            int Src_last_row = Src.Row, Src_last_col = Src.Column;
            int Dst_last_row = Dst.Row, Dst_last_col = Dst.Column;

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
            CopyValue2(ws_Src, ws_Dst, TestCase.NameDefinitionRow, 1, TestCase.NameDefinitionRow, Src_last_col );

            // Copy [Src_last_row,1] from tc to template
            CopyValue2(ws_Src, ws_Dst, Src_last_row, 1);

            // Copy the rest of data
            CopyValue2(ws_Src, ws_Dst, TestCase.DataBeginRow, 1, Src_last_row - 1, Src_last_col);

            return true;
        }
    }
}
