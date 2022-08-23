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
        static public Excel.Application OpenOridnaryExcel(string filename)
        {
            // Open excel (read-only)
            Excel.Application myBugExcel = new Excel.Application();
            Workbook working_book = myBugExcel.Workbooks.Open(filename, ReadOnly: true);
            myBugExcel.Visible = ExcelVisible;
            return myBugExcel;
        }

        static public Excel.Application OpenPreviousExcel(string filename)
        {
            // Open excel (read-only & corrupt-load)
            Excel.Application myBugExcel = new Excel.Application();
            //Workbook working_book = myBugExcel.Workbooks.Open(filename)
            //Workbook working_book = myBugExcel.Workbooks.Open(filename, ReadOnly: true, CorruptLoad: XlCorruptLoad.xlExtractData);
            myBugExcel.Workbooks.Open(filename, ReadOnly: true, CorruptLoad: XlCorruptLoad.xlExtractData);
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

        static public Dictionary<string, int> CreateTableColumnIndex(Worksheet bug_worksheet, int naming_row)
        {
            // Get the last (row,col) of excel
            Range rngLast = bug_worksheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            Dictionary<string, int> col_name_list = new Dictionary<string, int>();

            for (int col_index = 1; col_index <= rngLast.Column; col_index++)
            {
                Object cell_value2 = bug_worksheet.Cells[naming_row, col_index].Value2;
                if (cell_value2 == null) { continue; }
                col_name_list.Add(cell_value2.ToString(), col_index);
            }

            return col_name_list;
        }

    }
}
