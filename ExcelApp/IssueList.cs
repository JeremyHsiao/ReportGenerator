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
        // constant strings for worksheet used in this application.
        static public string SheetName = "general_report";
        static public int NameDefinitionRow = 4;
        static public int DataBeginRow = 5;
         // Key value
        static public string KeyPrefix = "BENSE";

        static public Dictionary<string, List<StyleString>> CreateBugListFromBugJiraFile(Worksheet bug_worksheet)
        {
            const string col_Key = "Key";
            const string col_Summary = "Summary";
            const string col_Severity = "Severity";
            const string col_RD_Comment = "Steps To Reproduce"; // To be updated 
            // const string col_RD_Comment = "Additional Information"; // To be updated
            //const int max_issue_no = 100000;
            Dictionary<string, List<StyleString>> bug_list = new Dictionary<string, List<StyleString>>();

            // Obtain column name listed on row 4 & its column index
            const int NameDefinitionRow = 4;
            Dictionary<string, int> col_name_list = ExcelAction.CreateTableColumnIndex(bug_worksheet, NameDefinitionRow);

            // Get the last (row,col) of excel
            Range rngLast = bug_worksheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

            // Collect bug info from row 5
            const int row_starting_key_index = 5;
            for (int row_index = row_starting_key_index; row_index <= rngLast.Row; row_index++)
            {
                List<StyleString> add_style_str = new List<StyleString>();
                string add_str, key_str, summary_str, severity_str, rd_commment_str;
                Object cell_value2;

                // Get Key string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_Key]].Value2;
                if (cell_value2 == null) { continue; }
                key_str = cell_value2.ToString();
                if (key_str.Contains(KeyPrefix) == false) { continue; }

                // Get Summary string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_Summary]].Value2;
                if (cell_value2 == null) { summary_str = ""; }
                else
                {
                    summary_str = cell_value2.ToString();
                    if (!summary_str.Substring(summary_str.Length - 1, 1).Equals(".")) { summary_str += "."; }  // make sure '.' at the end
                }

                // Get Severity string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_Severity]].Value2;
                if (cell_value2 == null) { severity_str = ""; }
                else { severity_str = cell_value2.ToString(); }

                // Get Description string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_RD_Comment]].Value2;
                if (cell_value2 == null) { rd_commment_str = ""; }
                else
                {
                    rd_commment_str = cell_value2.ToString();
                    // Remove strings after 1st '\n'
                    int new_line_pos = rd_commment_str.IndexOf('\n');
                    if (new_line_pos > 0)
                    {
                        rd_commment_str = rd_commment_str.Substring(0, new_line_pos);
                    }
                }

                // setup id/string pair
                StyleString str;

                add_str = key_str + summary_str + "(" + severity_str + ")";
                str = new StyleString(add_str, Color.Red);
                add_style_str.Add(str);
                if (rd_commment_str.CompareTo("") != 0)
                {
                    add_str = " --> " + rd_commment_str;
                    str = new StyleString(add_str, Color.Blue);
                    add_style_str.Add(str);
                }
                bug_list.Add(key_str, add_style_str);
            }
            return bug_list;
        }
    }
}
