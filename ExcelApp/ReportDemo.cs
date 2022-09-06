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

        // 
        // This demo open Test Case Excel and replace Issue ID on Linked Issue column with ID+Summary+Severity+RD_Comment
        //
        static public void WriteBacktoTCJiraExcel(String tclist_filename)
        {
            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename);

            if (status == ExcelAction.ExcelStatus.OK)
            {
                Dictionary<string, int> col_name_list = ExcelAction.CreateTestCaseColumnIndex();

                int key_col = col_name_list[TestCase.col_Key];
                int links_col = col_name_list[TestCase.col_LinkedIssue];
                // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
                for (int index = TestCase.DataBeginRow; index <= ExcelAction.GetTestCaseAllRange().Row; index++)
                {
                    // Make sure Key of TC contains KeyPrefix
                    String key = ExcelAction.GetTestCaseCellTrimmedString(index, key_col);
                    if (key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row

                    // If Links is not empty, extend bug key into long string with font settings
                    String links = ExcelAction.GetTestCaseCellTrimmedString(index, links_col);
                    if (links != "")
                    {
                        List<StyleString> str_list = StyleString.ExtendIssueDescription(links,global_issue_description_list);
                        ExcelAction.TestCase_WriteStyleString(index, links_col, str_list);
                    }
                }
                // auto-fit-height of column links
                ExcelAction.TestCase_AutoFit_Column(links_col);

                // Write to another filename with datetime
                string dest_filename = FileFunction.GenerateFilenameWithDateTime(tclist_filename);
                ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename);
            }
            // Always try to close at the end even there may be some error during operation
            ExcelAction.CloseTestCaseExcel();
        }


        // 
        // This demo open Summary Report Excel and write to Notes with all issues beloging to this test group (issue written in ID+Summary+Severity+RD_Comment)
        //
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
                            str_list = StyleString.ExtendIssueDescription(note, global_issue_description_list);
                            rng = result_worksheet.Cells[index, col_issue];
                            StyleString.WriteStyleString(ref rng, str_list);
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

        static public void ConsoleWarning(String function, int row)
        {
            Console.WriteLine("Warning: please check " + function + " at line " + row.ToString());
        }
        static public void ConsoleWarning(String function)
        {
            Console.WriteLine("Warning: please check " + function);
        }

        //
        // This Demo Identify Keyword on the excel and insert a column to list all issues containing that keyword
        //
        static int col_indentifier = 2;
        static int col_keyword = 3;
        static public bool KeywordIssueGenerationTask(string report_filename)
        {
            //
            // 1. Open Excel and find the sheet
            //

            String full_filename = FileFunction.GetFullPath(report_filename);
            String short_filename = Path.GetFileName(full_filename);
            String sheet_name = short_filename.Substring(0, short_filename.IndexOf("_"));

            if (!FileFunction.FileExists(full_filename))
            {
                ConsoleWarning("FileExists in KeywordIssueGenerationTask");
                return false;
            }

            Excel.Application myReportExcel = ExcelAction.OpenOridnaryExcel(full_filename, ReadOnly:false);
            if (myReportExcel == null)
            {
                ConsoleWarning("OpenOridnaryExcel in KeywordIssueGenerationTask");
                return false;
            }

            Worksheet result_worksheet = ExcelAction.Find_Worksheet(myReportExcel, sheet_name);
            if (result_worksheet == null)
            {
                ConsoleWarning("Find_Worksheet in KeywordIssueGenerationTask");
                return false;
            }

            //
            // 2. Find out Printable Area
            //

            String PrintArea = result_worksheet.PageSetup.PrintArea;
            Range rngPrintable = result_worksheet.Range[PrintArea];
            int row_print_area, column_print_area;
            // Data processing starting at "$A$1"
            // ending at Printable aread
            row_print_area = rngPrintable.Rows.Count;
            column_print_area = rngPrintable.Columns.Count;

            //
            // 3. Find out all keywords and create LUT (keyword,row_index)
            //    output:  LUT (keyword,row_index)
            //
            const int row_test_detail_start = 27;
            const String identifier_str = "Test Item";
            // Read report file for keyword & its row and store into keyword/row dictionary
            // Search keyword within printable area
            Dictionary<String, int> KeywordAtRow = new Dictionary<String, int>();
            for (int row_index = row_test_detail_start; row_index <= row_print_area; row_index++)
            {
                Object cell_obj = result_worksheet.Cells[row_index, col_indentifier].Value2;
                if(cell_obj==null) continue;
                String cell_text = cell_obj.ToString().Trim();
                if ((cell_text.Length>identifier_str.Length) &&
                    String.Equals(cell_text.Substring(0,identifier_str.Length), identifier_str, StringComparison.OrdinalIgnoreCase))
                {
                    cell_obj = result_worksheet.Cells[row_index, col_keyword].Value2;
                    if(cell_obj==null) { ConsoleWarning("Empty Keyword", row_index); continue;}
                    cell_text = cell_obj.ToString().Trim();
                    if (cell_text == "") { ConsoleWarning("Empty Keyword", row_index); continue; }
                    if(KeywordAtRow.ContainsKey(cell_text))
                    { ConsoleWarning("Duplicated Keyword", row_index); continue; }
                    KeywordAtRow.Add(cell_text, row_index);
                }
            }

            //
            // 4. Use keyword to find out all issues that contains keyword. 
            //    put issue_id into a string contains many id separated by a comma ','
            //    then store this issue_id into LUT (keyword,ids)
            //    output: LUT (keyword,id_list)
            //
            Dictionary<String, List<String>> KeywordIssueIDList = new Dictionary<String, List<String>>();
            foreach (String keyword in KeywordAtRow.Keys)
            {
                List<String> id_list = new List<String>();
                foreach (IssueList issue in global_issue_list)
                {
                    if (issue.Summary.Contains(keyword))
                    {
                        id_list.Add(issue.Key);
                    }
                }
                KeywordIssueIDList.Add(keyword, id_list);
            }

            //
            // 5. input:  LUT (keyword,id_list) + LUT (id,color_desription) (from GenerateIssueDescription())
            //    output: LUT (keyword,color_desription_list)
            //         
            //    using: id_list -> ExtendIssueDescription() -> color_description_list
            // This issue description list is needfed for keyword issue list
            global_issue_description_list = IssueList.GenerateIssueDescription(global_issue_list);

            // Go throught each keyword and turn id_list into color_description
            Dictionary<String, List<StyleString>> KeyWordIssueDescription = new Dictionary<String, List<StyleString>>();
            foreach (String keyword in KeywordAtRow.Keys)
            {
                List<String> id_list = KeywordIssueIDList[keyword];
                List<StyleString> issue_description = StyleString.ExtendIssueDescription(id_list, global_issue_description_list);
                KeyWordIssueDescription.Add(keyword, issue_description);
            }

            //
            // 6. input:  LUT (keyword,color_description_list) + LUT (id,row_index)
            //    output: write color_description_list at Excel(row_index,new_inserted_col outside printable area
            //         
            // Insert extra column just outside printable area.
            int insert_col = column_print_area + 1;
            result_worksheet.Columns[insert_col].Insert();

            foreach (String keyword in KeywordAtRow.Keys)
            {
                List<StyleString> issue_description = KeyWordIssueDescription[keyword];
                Range rng = result_worksheet.Cells[KeywordAtRow[keyword], insert_col];
                StyleString.WriteStyleString(ref rng, issue_description);
            }

            // Save as another file with yyyyMMddHHmmss
            string dest_filename = FileFunction.GenerateFilenameWithDateTime(full_filename);
            ExcelAction.SaveChangesAndCloseExcel(myReportExcel, dest_filename);

            return true;
        }

        //
        // This demo finds out Test-case whose status is fail but all linked issues are closed (other issues are hidden)
        //
        static String[] CloseStatusString = { IssueList.STR_CLOSE };
        static public void FindFailTCLinkedIssueAllClosed(String tclist_filename)
        {
            // protection-check
            if (global_issue_list.Count == 0)
            {
                ConsoleWarning("Bug list is empty");  return;
            }
            if (global_testcase_list.Count == 0)
            {
                ConsoleWarning("Test-case list is empty"); return;
            }

            // Prepare a list of key whose status is closed (waived treated as non-closed at the moment)
            List<String> ClosedIssueKey = new List<String>();
            foreach(IssueList issue in global_issue_list)
            {
                foreach (String str in CloseStatusString)
                {
                    if (issue.Status == str)
                    {
                        // if status is "close" or alike, add key into list and leave this loop
                        ClosedIssueKey.Add(issue.Key);
                        break;
                    }
                }
            }

            // Prepare several lists to separate TC into different groups
            List<String> tc_pass = new List<String>();                      // TC Status is Pass
            List<String> tc_none = new List<String>();                      // TC Status is None
            List<String> tc_fail_empty_link_issue = new List<String>();     // TC Status is Fail AND Links are empty
            List<String> tc_fail_some_nonclosed = new List<String>();       // TC Status is Fail AND Links have at least one non-closed issue
            List<String> tc_fail_all_closed = new List<String>();           // TC Status is Fail AND Links are all closed
            // looping all TC where links are not empty
            foreach (TestCase tc in global_testcase_list) // looping
            {
                if (tc.Status == TestCase.STR_PASS)
                {
                    tc_pass.Add(tc.Key);
                }
                else if (tc.Status == TestCase.STR_NONE)
                {
                    tc_none.Add(tc.Key);
                }
                else if (tc.Links.Trim() == "") // fail but empty linked issue
                {
                    tc_fail_empty_link_issue.Add(tc.Key);
                }
                else 
                {
                    List<String> LinkedIssueKey = TestCase.Convert_LinksString_To_ListOfString(tc.Links);
                    IEnumerable<String> LinkIssue_CloseIssue_intersect = ClosedIssueKey.Intersect(LinkedIssueKey);
                    if (LinkIssue_CloseIssue_intersect.Count() != LinkedIssueKey.Count())
                    {
                        // One ore more linked issue are not close (or close-alike), add into this list
                        tc_fail_some_nonclosed.Add(tc.Key);
                    }
                    else
                    {
                        tc_fail_all_closed.Add(tc.Key);
                    }
                }
            }

            // Start to hide rows unless this row belongs to tc_fail_all_closed

            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename);

            if (status == ExcelAction.ExcelStatus.OK)
            {
                Dictionary<string, int> col_name_list = ExcelAction.CreateTestCaseColumnIndex();
                int key_col = col_name_list[TestCase.col_Key];
                // Visit all rows to check if key belongs to tc_fail_all_closed

                for (int index = TestCase.DataBeginRow; index <= ExcelAction.GetTestCaseAllRange().Row; index++)
                {
                    // Make sure Key of TC contains KeyPrefix
                    String key = ExcelAction.GetTestCaseCellTrimmedString(index, key_col);
                    if (key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row

                    if (tc_fail_all_closed.Count == 0) { ExcelAction.TestCase_Hide_Row(index); }
                    else if (tc_fail_all_closed.IndexOf(key) < 0) { ExcelAction.TestCase_Hide_Row(index); }
                }

                // Write to another filename with datetime
                string dest_filename = FileFunction.GenerateFilenameWithDateTime(tclist_filename);
                ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename);
            }
            // Always try to close at the end even there may be some error during operation
            ExcelAction.CloseTestCaseExcel();
        }

    }
}
