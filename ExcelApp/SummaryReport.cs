using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReportApplication
{
    class SummaryReport
    {
        // 
        // This demo open Summary Report Excel and write to Notes with all issues beloging to this test group (issue written in ID+Summary+Severity+RD_Comment)
        //
        static public string sheet_Report_Result = "Result";
        static public void SaveIssueToSummaryReport(string report_filename)
        {
            // Re-arrange test-case list into dictionary of summary/links pair
            Dictionary<String, String> group_note_issue = new Dictionary<String, String>();
            foreach (TestCase tc in ReportGenerator.ReadGlobalTestCaseList())
            {
                String key = tc.Summary;
                //if (key != "")
                if (String.IsNullOrWhiteSpace(key) == false)
                {
                    group_note_issue.Add(key, tc.LinkedBug);
                }
            }

            Workbook wb_summary = ExcelAction.OpenExcelWorkbook(report_filename);
            if (wb_summary != null)
            {
                Worksheet result_worksheet = ExcelAction.Find_Worksheet(wb_summary, sheet_Report_Result);
                if (result_worksheet != null)
                {
                    //const int result_NameDefinitionRow = 5;
                    //const string col_Key = "TEST   ITEM";
                    //const string col_Links = "Links";
                    //Dictionary<string, int> result_col_name_list = CreateTableColumnIndex(result_worksheet, result_NameDefinitionRow);

                    // Get the last (row,col) of excel
                    Range rngLast = ExcelAction.GetWorksheetAllRange(result_worksheet);
                    const int col_group = 1, col_result = 2, col_issue = 3; // column "A" - "C"
                    const int row_result_starting = 6; // row starting from 6

                    int end_row = ExcelAction.Get_Range_RowNumber(rngLast);
                    for (int index = row_result_starting; index <= end_row; index++)
                    {
                        List<StyleString> str_list = new List<StyleString>();
                        String key, note;

                        // find out which test_group
                        key = ExcelAction.GetCellTrimmedString(result_worksheet, index, col_group);
                        if (String.IsNullOrWhiteSpace(key)) break; // if no value in test_group-->end of report

                        // goes to next row if Result is N/A
                        if (ExcelAction.GetCellTrimmedString(result_worksheet, index, col_result) == "N/A") continue;

                        // Get data to be filled into Note
                        // if key does not exist, Note will be empty string
                        if (!group_note_issue.TryGetValue(key, out note))
                        {
                            note = "";
                        }

                        if (String.IsNullOrWhiteSpace(note) == false)
                        {
                            // issue --> Fail
                            ExcelAction.SetCellValue(result_worksheet, index, col_result, "Fail");
                            // Fill "Note" 
                            List<Issue> issue_list = Issue.KeyStringToListOfIssue(note, ReportGenerator.ReadGlobalIssueList());
                            str_list = StyleString.BugList_To_SummaryPageFullIssueDescription(issue_list);
                            StyleString.WriteStyleString(result_worksheet, index, col_issue, str_list);
                        }
                        else
                        {
                            // no issue --> Pass
                            ExcelAction.SetCellValue(result_worksheet, index, col_result, "Pass");
                            ExcelAction.SetCellValue(result_worksheet, index, col_issue, "");
                        }

                        // auto-fit-height of current row
                        ExcelAction.AutoFit_Row(result_worksheet, index);
                    }

                    // Save as another file with yyyyMMddHHmmss
                    String dest_filename = Storage.GenerateFilenameWithDateTime(report_filename);
                    ExcelAction.CloseExcelWorkbook(wb_summary, SaveChanges: true, AsFilename: dest_filename);
                }
                else
                {
                    // worksheet not found, close immediately
                    ExcelAction.CloseExcelWorkbook(wb_summary, SaveChanges: false);
                }
            }
        }

    }
}
