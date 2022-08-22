using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ExcelReportApplication
{
    class IssueList
    {
        private String key;
        private String summary;
        private String severity;
        private String comment;

        public String Key   // property
        {
            get { return key; }   // get method
            set { key = value; }  // set method
        }

        public String Summary   // property
        {
            get { return summary; }   // get method
            set { summary = value; }  // set method
        }

        public String Severity   // property
        {
            get { return severity; }   // get method
            set { severity = value; }  // set method
        }

        public String Comment   // property
        {
            get { return comment; }   // get method
            set { comment = value; }  // set method
        }

        public const string col_Key = "Key";
        public const string col_Summary = "Summary";
        public const string col_Severity = "Severity";
        public const string col_RD_Comment = "Steps To Reproduce"; // To be updated 
        // public const string col_RD_Comment = "Additional Information"; // To be updated
        public IssueList()
        {
        }

        public IssueList(String key, String summary, String severity, String comment)
        {
            this.key = key; this.summary = summary; this.severity = severity; this.comment = comment;
        }

        // constant strings for worksheet used in this application.
        static public string SheetName = "general_report";
        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
         // Key value
        static public string KeyPrefix = "BENSE";

        static public List<IssueList> GenerateIssueList(string buglist_filename)
        {
            List<IssueList> ret_issue_list = new List<IssueList>();

            // Open excel (read-only & corrupt-load)
            Excel.Application myIssueExcel = ExcelAction.OpenPreviousExcel(buglist_filename);
            //Excel.Application myIssueExcel = OpenOridnaryExcel(buglist_filename);
            if (myIssueExcel != null)
            {
                Worksheet WorkingSheet = ExcelAction.Find_Worksheet(myIssueExcel, SheetName);
                if (WorkingSheet != null)
                {
                    Dictionary<string, int> col_name_list = ExcelAction.CreateTableColumnIndex(WorkingSheet, NameDefinitionRow);

                    // Get the last (row,col) of excel
                    Range rngLast = WorkingSheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                    // Visit all rows and add content of TestCase
                    for (int index = DataBeginRow; index <= rngLast.Row; index++)
                    {
                        Object cell_value2;
                        String key, summary, severity, comment;

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[IssueList.col_Key]].Value2;
                        key = (cell_value2 == null) ? "" : cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[IssueList.col_Summary]].Value2;
                        summary = (cell_value2 == null) ? "" : cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[IssueList.col_Severity]].Value2;
                        severity = (cell_value2 == null) ? "" : cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[IssueList.col_RD_Comment]].Value2;
                        comment = (cell_value2 == null) ? "" : cell_value2.ToString();

                        ret_issue_list.Add(new IssueList(key, summary, severity, comment));
                    }
                }
                ExcelAction.CloseExcelWithoutSaveChanges(myIssueExcel);
                myIssueExcel = null;
            }
            return ret_issue_list;
        }

        static public Color descrption_color_issue = Color.Red;
        static public Color descrption_color_comment = Color.Blue;
        static public Dictionary<string, List<StyleString>> CreateFullIssueDescription(List<IssueList> issuelist)
        {
            Dictionary<string, List<StyleString>> ret_list = new Dictionary<string, List<StyleString>>();

            foreach (IssueList issue in issuelist)
            {
                List<StyleString> value_style_str = new List<StyleString>();
                String key = issue.Key, rd_commment_str = issue.comment;

                if (key != "")
                {
                    String str = key + issue.Summary + "(" + issue.Severity + ")";
                    StyleString style_str = new StyleString(str, descrption_color_issue);
                    value_style_str.Add(style_str);
                    if (rd_commment_str != "")
                    {
                        str = " --> " + rd_commment_str;
                        style_str = new StyleString(str, descrption_color_comment);
                        value_style_str.Add(style_str);
                    }
                    ret_list.Add(key, value_style_str);
                }
            }
            return ret_list;
        }
    }
}
