using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace ExcelReportApplication
{
    static class ReportDemo
    {
        static public List<IssueList> global_issue_list = new List<IssueList>();
        static public Dictionary<string, List<StyleString>> global_issue_description_list = new Dictionary<string, List<StyleString>>();
        static public List<TestCase> global_testcase_list = new List<TestCase> ();

        static public List<StyleString> ExtendIssueDescription(string links_str, Dictionary<string, List<StyleString>> bug_list)
        {
            List<StyleString> extended_str = new List<StyleString>();

            // protection
            if ((links_str == null) || (bug_list == null)) return null;

            // Separate keys
            string[] separators = { "," };
            string[] issues = links_str.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            // replace key with full description and combine into one string
            foreach (string key in issues)
            {
                string trimmed_key = key.Trim();
                StyleString new_line_str = new StyleString("\n");
                if (bug_list.ContainsKey(trimmed_key))
                {
                    List<StyleString> bug_str = bug_list[trimmed_key]; 

                    foreach (StyleString style_str in bug_str)
                    {

                        extended_str.Add(style_str);
                    }
                }
                else
                {
                    StyleString def_str = new StyleString(trimmed_key);
                    extended_str.Add(def_str);
                }
                extended_str.Add(new_line_str);
            }
            if (extended_str.Count > 0) { extended_str.RemoveAt(extended_str.Count - 1); } // remove last '\n'
 
            return extended_str;
        }

        static public void WriteSytleString(ref Range input_range, List<StyleString> sytle_string_list)
        {
            // Fill the text into excel cell with default font settings.
            string txt_str = "";
            foreach (StyleString style_str in sytle_string_list)
            {
                txt_str += style_str.Text;
            }
            input_range.Value2 = txt_str;
            input_range.Characters.Font.Name = StyleString.default_font;
            input_range.Characters.Font.Size = StyleString.default_size;
            input_range.Characters.Font.Color = StyleString.default_color;
            input_range.Characters.Font.FontStyle = StyleString.default_fontstyle;

            // Change font settings when required for the string portion
            int chr_index = 1;
            foreach (StyleString style_str in sytle_string_list)
            {
                int len = style_str.Text.Length;
                if (style_str.FontPropertyChanged == true)
                {
                    input_range.get_Characters(chr_index, len).Font.Name = style_str.Font;
                    input_range.get_Characters(chr_index, len).Font.Color = style_str.Color;
                    input_range.get_Characters(chr_index, len).Font.Size = style_str.Size;
                    input_range.get_Characters(chr_index, len).Font.FontStyle = style_str.FontStyle;
                }
                chr_index += len;
            }
        }

        //WriteBacktoTCJiraExcel
        static public void WriteBacktoTCJiraExcel(String tclist_filename)
        {
            // Re-arrange test-case list into dictionary of key/links pair
            Dictionary<String, String> group_note_issue = new Dictionary<String, String>();
            foreach (TestCase tc in ReportDemo.global_testcase_list)
            {
                String key = tc.Key;
                if (key != "")
                {
                    group_note_issue.Add(key, tc.Links);
                }
            }

            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            Excel.Application myTCExcel = ExcelAction.OpenPreviousExcel(tclist_filename);
            if (myTCExcel != null)
            {
                Worksheet WorkingSheet = ExcelAction.Find_Worksheet(myTCExcel, TestCase.SheetName);
                if (WorkingSheet != null)
                {
                    Dictionary<string, int> col_name_list = ExcelAction.CreateTableColumnIndex(WorkingSheet, TestCase.NameDefinitionRow);

                    // Get the last (row,col) of excel
                    Range rngLast = WorkingSheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                    // Visit all rows and replace Bug-ID with long description of Bug.
                    for (int index = TestCase.DataBeginRow; index <= rngLast.Row; index++)
                    {
                        Object cell_value2;

                        // Make sure Key of TC is not empty
                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Key]].Value2;
                        if (cell_value2 == null) { break; }
                        String key = cell_value2.ToString();
                        if (key.Contains(TestCase.KeyPrefix) == false) { break; }

                        // If Links is not empty, extend bug key into long string with font settings
                        Range rng = WorkingSheet.Cells[index, col_name_list[TestCase.col_Links]];
                        cell_value2 = rng.Value2;
                        if (cell_value2 != null)
                        {
                            List<StyleString> str_list = ReportDemo.ExtendIssueDescription(group_note_issue[key],
                                                                            ReportDemo.global_issue_description_list);

                            ReportDemo.WriteSytleString(ref rng, str_list);
                        }
                    }
                    // auto-fit-height of column links
                    WorkingSheet.Columns[col_name_list[TestCase.col_Links]].AutoFit();

                    // Write to another filename with datetime
                    string dest_filename = FileFunction.GenerateFilenameWithDateTime(tclist_filename);
                    ExcelAction.SaveChangesAndCloseExcel(myTCExcel, dest_filename);
                }
                else
                {
                    // worksheet not found, close immediately
                    ExcelAction.CloseExcelWithoutSaveChanges(myTCExcel);
                }
                WorkingSheet = null;
                myTCExcel = null;
            }
        }

        static public string sheet_Report_Result = "Result";
        static public void SaveToReportTemplate(string report_filename)
        {
            // Re-arrange test-case list into dictionary of summary/links pair
            Dictionary<String, String> group_note_issue = new Dictionary<String, String>();
            foreach (TestCase tc in global_testcase_list)
            {
                String key = tc.Summary;
                if (key != "")
                {
                    group_note_issue.Add(key, tc.Links);
                }
            }

            //Excel.Application myReportExcel = ExcelAction.OpenPreviousExcel(report_filename);
            Excel.Application myReportExcel = ExcelAction.OpenOridnaryExcel(report_filename);
            if (myReportExcel != null)
            {
                Worksheet result_worksheet = ExcelAction.Find_Worksheet(myReportExcel, sheet_Report_Result);
                if (result_worksheet != null)
                {
                    //const int result_NameDefinitionRow = 5;
                    //const string col_Key = "TEST   ITEM";
                    //const string col_Links = "Links";
                    //Dictionary<string, int> result_col_name_list = CreateTableColumnIndex(result_worksheet, result_NameDefinitionRow);

                    // Get the last (row,col) of excel
                    Range rngLast = result_worksheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                    const int col_group = 1, col_result = 2, col_issue = 3; // column "A" - "C"
                    const int row_result_starting = 6; // row starting from 6

                    for (int index = row_result_starting; index <= rngLast.Row; index++)
                    {
                        Range rng;
                        Object cell_value2; 
                        List<StyleString> str_list = new List<StyleString>();
                        String key, note;

                        // find out which test_group
                        rng = result_worksheet.Cells[index, col_group];
                        cell_value2 = rng.Value2;
                        if (cell_value2 == null) { break; } // if no value in test_group-->end of report
                        key = cell_value2.ToString();

                        // goes to next row if Result is N/A
                        rng = result_worksheet.Cells[index, col_result];
                        if (rng.Value2.ToString().Trim() == "N/A") { continue; } // goes to next row if N/A
 
                        // Get data to be filled into Note
                        // if key does not exist, Note will be empty string
                        if (!group_note_issue.TryGetValue(key, out note))
                        {
                            note = "";
                        }

                        if (note!="")
                        {
                            rng = result_worksheet.Cells[index, col_result];
                            rng.Value2 = "Fail";
                            // Fill "Note" 
                            str_list = ExtendIssueDescription(note, global_issue_description_list);
                            rng = result_worksheet.Cells[index, col_issue];
                            WriteSytleString(ref rng, str_list);
                        }
                        else
                        {
                            // no issue --> pass
                            rng = result_worksheet.Cells[index, col_result];
                            rng.Value2 = "Pass";
                            rng = result_worksheet.Cells[index, col_issue];
                            rng.Value2 = "";
                        }

                        // auto-fit-height of current row
                        rng.Rows.AutoFit();
                     }

                    // Save as another file with yyyyMMddHHmmss
                    string dest_filename = FileFunction.GenerateFilenameWithDateTime(report_filename);
                    ExcelAction.SaveChangesAndCloseExcel(myReportExcel, dest_filename);
                }
                else
                {
                    // worksheet not found, close immediately
                    ExcelAction.CloseExcelWithoutSaveChanges(myReportExcel);
                }
            }
        }
    }
}
