using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReportApplication
{
    public class TestCase
    {
        private String key;
        private String group;
        private String summary;
        private String status;
        private String links;

        public String Key   // property
        {
            get { return key; }   // get method
            set { key = value; }  // set method
        }

        public String Group   // property
        {
            get { return group; }   // get method
            set { group = value; }  // set method
        }

        public String Summary   // property
        {
            get { return summary; }   // get method
            set { summary = value; }  // set method
        }

        public String Status   // property
        {
            get { return status; }   // get method
            set { status = value; }  // set method
        }

        public String Links   // property
        {
            get { return links; }   // get method
            set { links = value; }  // set method
        }

        public const string col_Key = "Key";
        public const string col_Group = "Test Group";
        public const string col_Summary = "Summary";
        public const string col_Status = "Status";
        public const string col_Links = "Linked Issues";
        public TestCase()
        {
        }

        public TestCase(String key, String group, String summary, String status, String links)
        {
            this.key = key; this.group = group; this.summary = summary; this.status = status; this.links = links;
        }

        static private String[] separators = { "," };
        static public List<String> Convert_LinksString_To_ListOfString(String links)
        {
            // protection
            if ((links == null)||(links =="")) return new List<String>();   // return empty new object
            // Separate keys into string[]
            String[] issues = links.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            if (issues == null) return null;
            // string[] to List<String> and return
            return issues.ToList();
        }

        static public String Convert_ListOfString_To_LinkString(List<String> list)
        {
            // protection
            if (list == null) return "";
            if (list.Count == 0) return "";
            String ret = "";
            foreach (String str in list)
            {
                ret += " " + str + ",";
            }
            ret.Remove(ret.Length - 1); // remove last ","
            ret.Trim(); // remove " " at beginning.
            return ret;
        }

        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
        static public string SheetName = "general_report";
        static public string KeyPrefix = "TCBEN";

        static public List<TestCase> GenerateTestCaseList(string tclist_filename)
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

            // Open excel (read-only & corrupt-load)
            Excel.Application myTCExcel = ExcelAction.OpenPreviousExcel(tclist_filename);
            //Excel.Application myTCExcel = OpenOridnaryExcel(tclist_filename);
            if (myTCExcel != null)
            {
                Worksheet WorkingSheet = ExcelAction.Find_Worksheet(myTCExcel, SheetName);
                if (WorkingSheet != null)
                {
                    Dictionary<string, int> col_name_list = ExcelAction.CreateTableColumnIndex(WorkingSheet, NameDefinitionRow);

                    // Get the last (row,col) of excel
                    Range rngLast = WorkingSheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                    // Visit all rows and add content of TestCase
                    for (int index = DataBeginRow; index <= rngLast.Row; index++)
                    {
                        Object cell_value2;
                        String key, group, summary, status, links;

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Key]].Value2;
                        key = (cell_value2 == null) ? "" : cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Group]].Value2;
                        group = (cell_value2 == null) ? "" : cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Summary]].Value2;
                        summary = (cell_value2 == null) ? "" : cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Status]].Value2;
                        status = (cell_value2 == null) ? "" : cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Links]].Value2;
                        links = (cell_value2 == null) ? "" : cell_value2.ToString();

                        ret_tc_list.Add(new TestCase(key, group, summary, status, links));
                    }
                }
                ExcelAction.CloseExcelWithoutSaveChanges(myTCExcel);
                myTCExcel = null;
            }
            return ret_tc_list;
        }
    }
}
