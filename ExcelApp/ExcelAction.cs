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
        static public bool ExcelVisible = true;

        // Open existing excel
        static public Excel.Application OpenOridnaryExcel(string filename, bool ReadOnly = true)
        {
            // Open excel (read-only)
            Excel.Application myBugExcel = new Excel.Application();
            Workbook working_book = myBugExcel.Workbooks.Open(filename, ReadOnly: ReadOnly);
            myBugExcel.Visible = ExcelVisible;
            return myBugExcel;
        }

        static public Excel.Application OpenPreviousExcel(string filename, bool ReadOnly = true)
        {
            // Open excel (read-only & corrupt-load)
            Excel.Application myBugExcel = new Excel.Application();
            //Workbook working_book = myBugExcel.Workbooks.Open(filename)
            //Workbook working_book = myBugExcel.Workbooks.Open(filename, ReadOnly: true, CorruptLoad: XlCorruptLoad.xlExtractData);
            myBugExcel.Workbooks.Open(filename, ReadOnly: ReadOnly, CorruptLoad: XlCorruptLoad.xlExtractData);
            myBugExcel.Visible = ExcelVisible;
            return myBugExcel;
        }

        static public void CloseExcelWithoutSaveChanges(Excel.Application myExcel)
        {
            myExcel.ActiveWorkbook.Close(SaveChanges: false);
            myExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            GC.Collect();
        }

        static public void SaveChangesAndCloseExcel(Excel.Application myExcel)
        {
            myExcel.DisplayAlerts = false;
            myExcel.ActiveWorkbook.Close(SaveChanges: true);
            myExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            GC.Collect();
        }

        static public void SaveChangesAndCloseExcel(Excel.Application myExcel, String filename)
        {
            myExcel.ActiveWorkbook.Close(SaveChanges: true, Filename: filename);
            myExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            GC.Collect();
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

        static public void Hide_Row(Worksheet ws, int row, int count = 1)
        {
            Range hiddenRange = ws.Range[ws.Cells[row, 1], ws.Cells[row + count - 1, GetWorksheetAllRange(ws).Column]];
            //     var hiddenRange = yourWorksheet.Range[yourWorksheet.Cells[firstRowToHide, firstColToHide], yourWorksheet.Cells[lastRowToHide, lastColToHide]];
            hiddenRange.EntireRow.Hidden = true;
        }

        static public Dictionary<string, int> CreateTableColumnIndex(Worksheet ws, int naming_row)
        {
            Dictionary<string, int> col_name_list = new Dictionary<string, int>();

            int col_end = GetWorksheetAllRange(ws).Column;
            for (int col_index = 1; col_index <= col_end; col_index++)
            {
                Object cell_value2 = ws.Cells[naming_row, col_index].Value2;
                if (cell_value2 == null) { continue; }
                col_name_list.Add(cell_value2.ToString(), col_index);
            }

            return col_name_list;
        }

        // Code for operations on specific Excel File

        public enum ExcelStatus  
        {
            OK = 0,
            INIT_STATE,
            ERR_OpenIssueListExcel_OpenPreviousExcel,
            ERR_OpenIssueListExcel_Find_Worksheet,
            ERR_OpenIssueListExcel_CloseExcelWithoutSaveChanges,
            ERR_OpenTestCaseExcel_OpenPreviousExcel,
            ERR_OpenTestCaseExcel_Find_Worksheet,
            ERR_OpenTestCaseExcel_CloseExcelWithoutSaveChanges,
            ERR_CloseIssueListExcel_close_null,
            ERR_CloseTestCaseExcel_close_null,
            ERR_SaveChangesAndCloseIssueListExcel_close_null,
            ERR_SaveChangesAndCloseTestCaseExcel_close_null,
            ERR_NOT_DEFINED,
            EX_OpenIssueListWorksheet,
            EX_CloseIssueListWorksheet,
            EX_SaveChangesAndCloseIssueListExcel,
            EX_OpenTestCaseWorksheet,
            EX_CloseTestCaseWorksheet,
            EX_SaveChangesAndCloseTestCaseExcel,
            MAX_NO
        };

        static private Excel.Application IssueListExcel;
        static private Workbook book_issuelist;
        static private Worksheet ws_issuelist;
        static private Excel.Application TestCaseExcel;
        static private Workbook book_testcase;
        static private Worksheet ws_testcase;
        static private Excel.Application TestCaseTemplateExcel;
        static private Workbook book_tc_template;
        static private Worksheet ws_tc_template;

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
            return CreateTableColumnIndex(ws_issuelist, IssueList.NameDefinitionRow);
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
                // Open excel (read-only & corrupt-load)
                Excel.Application myIssueExcel = ExcelAction.OpenPreviousExcel(buglist_filename);

                if (myIssueExcel == null)
                {
                    return ExcelStatus.ERR_OpenIssueListExcel_OpenPreviousExcel;
                }

                Worksheet ws_buglist = ExcelAction.Find_Worksheet(myIssueExcel, IssueList.SheetName);
                if (ws_buglist == null)
                {
                    return ExcelStatus.ERR_OpenIssueListExcel_Find_Worksheet;
                }
                else
                {
                    IssueListExcel = myIssueExcel;
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
                if (IssueListExcel == null)
                {
                    return ExcelStatus.ERR_CloseIssueListExcel_close_null;
                }
                ExcelAction.CloseExcelWithoutSaveChanges(IssueListExcel);
                ws_issuelist = null;
                IssueListExcel = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_issuelist = null;
                IssueListExcel = null;
                return ExcelStatus.EX_CloseIssueListWorksheet;
            }
        }

        static public ExcelStatus SaveChangesAndCloseIssueListExcel(String dest_filename)
        {
            try
            {
                if (IssueListExcel == null)
                {
                    return ExcelStatus.ERR_SaveChangesAndCloseIssueListExcel_close_null;
                }
                ExcelAction.SaveChangesAndCloseExcel(IssueListExcel, dest_filename);
                ws_issuelist = null;
                IssueListExcel = null;
                return ExcelStatus.OK;
            }
            catch
            {
                ws_issuelist = null;
                IssueListExcel = null;
                return ExcelStatus.EX_SaveChangesAndCloseIssueListExcel;
            }
        }

        // Excel Open/Close/Save for Test Case Excel
        static public ExcelStatus OpenTestCaseExcel(String tclist_filename, bool IsTemplate = false)
        {
            try
            {
                Excel.Application myTCExcel;
                if (IsTemplate == false)
                {
                    // Open excel (read-only & corrupt-load)
                    myTCExcel = ExcelAction.OpenPreviousExcel(tclist_filename);
                }
                else
                {
                    myTCExcel = ExcelAction.OpenOridnaryExcel(tclist_filename);
                }

                if (myTCExcel == null)
                {
                    return ExcelStatus.ERR_OpenTestCaseExcel_OpenPreviousExcel;
                }

                Worksheet ws_tclist = ExcelAction.Find_Worksheet(myTCExcel, TestCase.SheetName);
                if (ws_tclist == null)
                {
                    return ExcelStatus.ERR_OpenTestCaseExcel_Find_Worksheet;
                }
                else
                {
                    if (IsTemplate == false)
                    {
                        TestCaseExcel = myTCExcel;
                        ws_testcase = ws_tclist;
                    }
                    else
                    {
                        TestCaseTemplateExcel = myTCExcel;
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
                    if (TestCaseExcel == null)
                    {
                        return ExcelStatus.ERR_CloseTestCaseExcel_close_null;
                    }
                    ExcelAction.CloseExcelWithoutSaveChanges(TestCaseExcel);
                    ws_testcase = null;
                    TestCaseExcel = null;
                }
                else
                {
                    if (TestCaseTemplateExcel == null)
                    {
                        return ExcelStatus.ERR_CloseTestCaseExcel_close_null;
                    }
                    ExcelAction.CloseExcelWithoutSaveChanges(TestCaseTemplateExcel);
                    ws_tc_template = null;
                    TestCaseTemplateExcel = null;
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
                    if (TestCaseExcel == null)
                    {
                        return ExcelStatus.ERR_SaveChangesAndCloseTestCaseExcel_close_null;
                    }
                    ExcelAction.SaveChangesAndCloseExcel(TestCaseExcel, dest_filename);
                    ws_testcase = null;
                    TestCaseExcel = null;
                }
                else
                {
                    if (TestCaseTemplateExcel == null)
                    {
                        return ExcelStatus.ERR_SaveChangesAndCloseTestCaseExcel_close_null;
                    }
                    ExcelAction.SaveChangesAndCloseExcel(TestCaseTemplateExcel, dest_filename);
                    ws_tc_template = null;
                    TestCaseTemplateExcel = null;
                }
                return ExcelStatus.OK;
            }
            catch
            {
                return ExcelStatus.EX_SaveChangesAndCloseTestCaseExcel;
            }
        }

        static public bool CopyTestCaseIntoTemplate()
        {
            // Protection
            if (ws_testcase == null) { return false; }
            if (ws_tc_template == null) { return false; }

            Range Src = GetWorksheetAllRange(ws_testcase);
            Range Dst = GetWorksheetAllRange(ws_tc_template);
            int Src_last_row = Src.Row, Src_last_col = Src.Column;
            int Dst_last_row = Dst.Row, Dst_last_col = Dst.Column;

            // Make template (destination) row count == TestCase (source) row count
            if (Src_last_row > Dst_last_row)
            {
                // Insert row into template file
                int rows_to_insert = Src_last_row - Dst_last_row;
                do
                {
                    ws_tc_template.Rows[TestCase.DataBeginRow + 1].Insert();
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

            // Copy [3,1] from tc to template
            Src = ws_testcase.Cells[3, 1];
            Dst = ws_tc_template.Cells[3, 1];
            Dst.Value2 = Src.Value2;

            // Copy row 4 (Column Name) from tc to template
            Src = ws_testcase.Range[ws_testcase.Cells[TestCase.NameDefinitionRow, 1], ws_testcase.Cells[TestCase.NameDefinitionRow, Src_last_col]];
            Dst = ws_tc_template.Range[ws_tc_template.Cells[TestCase.NameDefinitionRow, 1], ws_tc_template.Cells[TestCase.NameDefinitionRow, Src_last_col]];
            Dst.Value2 = Src.Value2;

            // Copy [Src_last_row,1] from tc to template
            Src = ws_testcase.Cells[Src_last_row, 1];
            Dst = ws_tc_template.Cells[Src_last_row, 1];
            Dst.Value2 = Src.Value2;

            // Copy the rest of data
            Src = ws_testcase.Range[ws_testcase.Cells[TestCase.DataBeginRow, 1], ws_testcase.Cells[Src_last_row - 1, Src_last_col]];
            Dst = ws_tc_template.Range[ws_tc_template.Cells[TestCase.DataBeginRow, 1], ws_tc_template.Cells[Src_last_row - 1, Src_last_col]];
            Dst.Value2 = Src.Value2;

            return true;
        }
    }
}
