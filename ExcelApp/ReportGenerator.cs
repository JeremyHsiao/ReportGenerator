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
    static class ReportGenerator
    {
        static public List<IssueList> global_issue_list = new List<IssueList>();
        static public Dictionary<string, List<StyleString>> global_full_issue_description_list = new Dictionary<string, List<StyleString>>();
        static public Dictionary<string, List<StyleString>> global_issue_description_list = new Dictionary<string, List<StyleString>>();
        static public List<TestCase> global_testcase_list = new List<TestCase>();

        public enum ReportType {
            FullIssueDescription_TC = 0,
            FullIssueDescription_Summary,
            StandardTestReportCreation,
            KeywordIssue_Report,
            TC_Likely_Passed
        }

        public static int ReportTypeToInt(ReportType type)
        {
            return (int)type; 
        }

        public static ReportType ReportTypeFromInt(int int_type)
        {
            return (ReportType)int_type;
        }

        public static int ReportTypeCount = Enum.GetNames(typeof(ReportType)).Length;

        public static String GetReportName(ReportType type)
        {
            return GetReportName(ReportTypeToInt(type));
        }

        public static String GetReportName(int type_index)
        {
            return ReportName[type_index];
        }

        public static List<String> ReportNameToList()
        {
            return ReportName.ToList();
        }

        private static String[] ReportName = new String[] 
        {
            "1.Issue Description for TC",
            "2.Issue Description for Summary",
            "3.Standard Test Report Creator",
            "4.Keyword Issue to Report",
            "5.TC likely Pass"
        };

        private static String[][] ReportDescription = new String[][] 
        {
            new String[] 
            {
                "Issue Description for TC Report", 
                "Input:",  "  Issue List + Test Case + Template (for Test case output)",
                "Output:", "  Test Case in the format of template file with linked issue in full description",
            },
            new String[] 
            {
                "Issue Description for Summary Report", 
                "Input:",  "  Issue List + Test Case + Template (for Summary Report)",
                "Output:", "  Summary in the format of template file with linked issue in full description",
            },
            new String[] 
            {
                "Create file structure of Standard Test Report according to user's selection (Do or Not)", 
                "Input:",  "  Main Test Report File",
                "Output:", "  Directory structure and report files under directories",
            },
            new String[] 
            {
                "Keyword Issue to Report", 
                "Input:",  "  Test Plan/Report with Keyword",
                "Output:", "  Test Plan/Report with keyword issue list inserted on O-column",
            },
            new String[] 
            {
                "Test case status is Fail but its linked issues are closed", 
                "Input:",  "  Issue List + Test Case + Template (for Test case output)",
                "Output:", "  Test Case whose linked issues are closed (other TC are hidden)",
            },
        };

        public static String GetReportDescription(int type_index)
        {
            String ret_str = "";
            foreach (String str in ReportDescription[type_index])
            {
                ret_str += str + "\r\n";
            }
            return ret_str;
        }

        public static String GetReportDescription(ReportType type)
        {
            return GetReportDescription(ReportTypeToInt(type));
        }


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
                        List<StyleString> str_list = StyleString.ExtendIssueDescription(links, global_full_issue_description_list);
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
        static public void SaveIssueToSummaryReport(string report_filename)
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

                    int end_row = rngLast.Row;
                    for (int index = row_result_starting; index <= end_row; index++)
                    {
                        List<StyleString> str_list = new List<StyleString>();
                        String key, note;

                        // find out which test_group
                        key = ExcelAction.GetCellTrimmedString(result_worksheet, index, col_group);
                        if (key == "") break; // if no value in test_group-->end of report

                        // goes to next row if Result is N/A
                        if (ExcelAction.GetCellTrimmedString(result_worksheet,index, col_result) == "N/A") continue;
 
                        // Get data to be filled into Note
                        // if key does not exist, Note will be empty string
                        if (!group_note_issue.TryGetValue(key, out note))
                        {
                            note = "";
                        }

                        if (note!="")
                        {
                            // issue --> Fail
                            ExcelAction.SetCellValue(result_worksheet, index, col_result, "Fail");
                            // Fill "Note" 
                            str_list = StyleString.ExtendIssueDescription(note, global_full_issue_description_list);
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
                    String dest_filename = FileFunction.GenerateFilenameWithDateTime(report_filename);
                    ExcelAction.CloseExcelWorkbook(wb_summary, SaveChanges: true, AsFilename: dest_filename);
                }
                else
                {
                    // worksheet not found, close immediately
                    ExcelAction.CloseExcelWorkbook(wb_summary, SaveChanges: false);
                }
            }
        }

        static private void ConsoleWarning(String function, int row)
        {
            Console.WriteLine("Warning: please check " + function + " at line " + row.ToString());
        }
        static private void ConsoleWarning(String function)
        {
            Console.WriteLine("Warning: please check " + function);
        }

        //
        // This demo finds out Test-case whose status is fail but all linked issues are closed (other issues are hidden)
        //
        static String[] CloseStatusString = { IssueList.STR_CLOSE };
        static public void FindFailTCLinkedIssueAllClosed(String tclist_filename, String template_filename)
        {
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

            // Open original excel (read-only & corrupt-load) and write to template file with another filename when closed

            // 1. open test case (as report source)
            ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                // TBD: what to do if cannot open template file
                ExcelAction.CloseTestCaseExcel();
            }
            else
            {

                // 2. open test case template
                status = ExcelAction.OpenTestCaseExcel(template_filename, IsTemplate: true);
                if (status != ExcelAction.ExcelStatus.OK)
                {
                    // TBD: what to do if cannot open template file
                    ExcelAction.CloseTestCaseExcel(IsTemplate: true);
                    ExcelAction.CloseTestCaseExcel();
                }
                else
                {
                    // 3. Copy test case data into template excel -- both will have the same row/col and (almost) same data
                    ExcelAction.CopyTestCaseIntoTemplate();
                    ExcelAction.CloseTestCaseExcel();

                    // 4. Excel processing on template excel
                    Dictionary<string, int> col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate:true);
                    int DataEndRow = ExcelAction.GetTestCaseAllRange(IsTemplate: true).Row;
                    int key_col = col_name_list[TestCase.col_Key];

                    // Visit all rows to check if key belongs to tc_fail_all_closed
                    int hide_row_start = 0, hide_row_count = 0;
                    for (int index = TestCase.DataBeginRow; index <= DataEndRow; index++)
                    {
                        // Make sure Key of TC contains KeyPrefix
                        String key = ExcelAction.GetTestCaseCellTrimmedString(index, key_col, IsTemplate: true);
                        if (key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row

                        bool blToHide = false;
                        if (tc_fail_all_closed.Count == 0) { blToHide = true; }
                        else if (tc_fail_all_closed.IndexOf(key) < 0) { blToHide = true; }
                        if (blToHide)
                        {
                            if (hide_row_start <= 0)
                            {
                                hide_row_start = index;
                            }
                            hide_row_count++;
                        }
                        else
                        {
                            // This row not to be hidden --> so hide all previous to-be-hidden rows
                            ExcelAction.TestCase_Hide_Row(hide_row_start, hide_row_count, IsTemplate:true);
                            hide_row_start = hide_row_count = 0;
                        }
                    }
                    // Hide allnot-hidden-yet rows
                    if ((hide_row_start > 0) && (hide_row_count > 0))
                    {
                        ExcelAction.TestCase_Hide_Row(hide_row_start, hide_row_count, IsTemplate: true);
                        hide_row_start = hide_row_count = 0;
                    }

                    // Save Template file as another filename (testcase filename with datetime & as .xlsx)
                    string dest_filename = FileFunction.GenerateFilenameWithDateTime(tclist_filename, ".xlsx");
                    ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename, IsTemplate: true);
                }
            }
        }

    }
}
