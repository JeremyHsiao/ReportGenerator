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

        static public void AutoFit_Column(Worksheet ws, int col)
        {
            ws.Columns[col].AutoFit();
        }


        static public Dictionary<string, int> CreateTableColumnIndex(Worksheet ws, int naming_row)
        {
            Dictionary<string, int> col_name_list = new Dictionary<string, int>();

            for (int col_index = 1; col_index <= GetWorksheetAllRange(ws).Column; col_index++)
            {
                Object cell_value2 = ws.Cells[naming_row, col_index].Value2;
                if (cell_value2 == null) { continue; }
                col_name_list.Add(cell_value2.ToString(), col_index);
            }

            return col_name_list;
        }
        
        public enum ExcelStatus  
        {
            OK = 0,
            INIT_STATE,
            ERR_OpenPreviousExcel,
            ERR_Find_Worksheet,
            ERR_CloseExcelWithoutSaveChanges,
            ERR_CloseTestCaseExcel_close_null,
            ERR_SaveChangesAndCloseExcel_close_null,
            ERR_NOT_DEFINED,
            EX_OpenTestCaseWorksheet,
            EX_CloseTestCaseWorksheet,
            MAX_NO
        };

        static private Excel.Application TestCaseExcel;
        static private Worksheet ws_testcase;

        static public Worksheet GetTestCaseWorksheet()
        {
            return ws_testcase;
        }

        static public Range GetTestCaseAllRange()
        {
            return ws_testcase.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
        }

        static public Object GetTestCaseCell(int row, int col)
        {
            return ws_testcase.Cells[row, col].Value2;
        }

        static public String GetTestCaseCellTrimmedString(int row, int col)
        {
            Object cell_value2 = GetTestCaseCell(row, col);
            if (cell_value2 == null) { return ""; }
            return cell_value2.ToString();
        }

        static public void TestCase_AutoFit_Column(int col)
        {
            AutoFit_Column(ws_testcase, col);
        }

        static public void TestCase_WriteStyleString(int row, int col, List<StyleString> sytle_string_list)
        {
            StyleString.WriteStyleString(ws_testcase, row, col, sytle_string_list);
        }

        static public Dictionary<string, int> CreateTestCaseColumnIndex()
        {
            return CreateTableColumnIndex(ws_testcase, TestCase.NameDefinitionRow);
        }

        static public ExcelStatus OpenTestCaseExcel(String tclist_filename)
        {
            try
            {
                // Open excel (read-only & corrupt-load)
                Excel.Application myTCExcel = ExcelAction.OpenPreviousExcel(tclist_filename);

                if (myTCExcel == null)
                {
                    return ExcelStatus.ERR_OpenPreviousExcel;
                }

                Worksheet ws_tclist = ExcelAction.Find_Worksheet(myTCExcel, TestCase.SheetName);
                if (ws_tclist == null)
                {
                    return ExcelStatus.ERR_Find_Worksheet;
                }
                else
                {
                    TestCaseExcel = myTCExcel;
                    ws_testcase = ws_tclist;
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

        static public ExcelStatus CloseTestCaseExcel()
        {
            try
            {
                if (TestCaseExcel == null)
                {
                    return ExcelStatus.ERR_CloseTestCaseExcel_close_null;
                }
                ExcelAction.CloseExcelWithoutSaveChanges(TestCaseExcel);
                TestCaseExcel = null;
                return ExcelStatus.OK;
            }
            catch
            {
                return ExcelStatus.EX_CloseTestCaseWorksheet;
            }
        }

        static public ExcelStatus SaveChangesAndCloseExcel(String dest_filename)
        {
            try
            {
                if (TestCaseExcel == null)
                {
                    return ExcelStatus.ERR_SaveChangesAndCloseExcel_close_null;
                }
                ExcelAction.SaveChangesAndCloseExcel(TestCaseExcel, dest_filename);
                TestCaseExcel = null;
                return ExcelStatus.OK;
            }
            catch
            {
                return ExcelStatus.EX_CloseTestCaseWorksheet;
            }
        }
    }
}
