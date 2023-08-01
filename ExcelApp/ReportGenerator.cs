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
        static public List<Issue> global_issue_list = new List<Issue>();
        static public Dictionary<string, List<StyleString>> global_full_issue_description_list = new Dictionary<string, List<StyleString>>();
        static public Dictionary<string, List<StyleString>> global_issue_description_list = new Dictionary<string, List<StyleString>>();
        static public Dictionary<string, List<StyleString>> global_issue_description_list_severity = new Dictionary<string, List<StyleString>>();
        static public List<TestCase> global_testcase_list = new List<TestCase>();
        static public List<String> fileter_status_list = new List<String>();
        static public List<ReportFileRecord> excel_not_report_log = new List<ReportFileRecord>();

        private static List<String> GetGlobalIssueKey(List<Issue> issue_list)
        {
            List<String> key_list = new List<String>();
            foreach (Issue issue in issue_list)
            {
                key_list.Add(issue.Key);
            }
            return key_list;
        }

        // 
        // This demo open Test Case Excel and replace Issue ID on Linked Issue column with ID+Summary+Severity+RD_Comment
        //
        //static public void WriteBacktoTCJiraExcel(String tclist_filename)
        //{
        //    // Open original excel (read-only & corrupt-load) and write to another filename when closed
        //    ExcelAction.ExcelStatus status = ExcelAction.OpenTestCaseExcel(tclist_filename);

        //    if (status == ExcelAction.ExcelStatus.OK)
        //    {
        //        Dictionary<string, int> col_name_list = ExcelAction.CreateTestCaseColumnIndex();

        //        int key_col = col_name_list[TestCase.col_Key];
        //        int links_col = col_name_list[TestCase.col_LinkedIssue];
        //        // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
        //        for (int index = TestCase.DataBeginRow; index <= ExcelAction.GetTestCaseAllRange().Row; index++)
        //        {
        //            // Make sure Key of TC contains KeyPrefix
        //            String key = ExcelAction.GetTestCaseCellTrimmedString(index, key_col);
        //            if (key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row

        //            // If Links is not empty, extend bug key into long string with font settings
        //            String links = ExcelAction.GetTestCaseCellTrimmedString(index, links_col);
        //            if (links != "")
        //            {
        //                List<StyleString> str_list = StyleString.ExtendIssueDescription(links, global_full_issue_description_list);
        //                ExcelAction.TestCase_WriteStyleString(index, links_col, str_list);
        //            }
        //        }
        //        // auto-fit-height of column links
        //        ExcelAction.TestCase_AutoFit_Column(links_col);

        //        // Write to another filename with datetime
        //        string dest_filename = FileFunction.GenerateFilenameWithDateTime(tclist_filename);
        //        ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename);
        //    }
        //    // Always try to close at the end even there may be some error during operation
        //    ExcelAction.CloseTestCaseExcel();
        //}

        // This version open Test Case Excel and copy content into template file and replace Issue ID on Linked Issue column with ID+Summary+Severity+RD_Comment
        static public void WriteBacktoTCJiraExcelV2(String tclist_filename, String template_filename, String judgement_report_dir = "")
        {
            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            ExcelAction.ExcelStatus status;

            status = ExcelAction.OpenTestCaseExcel(tclist_filename);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                ExcelAction.CloseTestCaseExcel();
                return; // to-be-checked if here
            }

            // 2. open test case template
            status = ExcelAction.OpenTestCaseExcel(template_filename, IsTemplate: true);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                ExcelAction.CloseTestCaseExcel();
                return; // to-be-checked if here
            }

            // 2.1 Get report_list under judgement_report_dir
            Dictionary<String, String> report_list = new Dictionary<String, String>();
            if (judgement_report_dir != "")
            {
                List<String> file_list = Storage.ListFilesUnderDirectory(judgement_report_dir);
                foreach (String name in file_list)
                {
                    // File existing check protection (it is better also checked and giving warning before entering this function)
                    if (Storage.FileExists(name) == false)
                        continue; // no warning here, simply skip this file.

                    String full_filename = Storage.GetFullPath(name);
                    //String short_filename = Storage.GetFileName(full_filename);
                    //String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                    //String sheet_name = sp_str[0];
                    String sheet_name = TestPlan.GetSheetNameAccordingToFilename(name);
                    try
                    {
                        report_list.Add(sheet_name, full_filename);
                    }
                    catch (ArgumentException)
                    {
                        Console.WriteLine("Sheet name:" + sheet_name + " already exists.");
                    }

                }
            }

            // 3. Copy test case data into template excel -- both will have the same row/col and (almost) same data
            ExcelAction.CopyTestCaseIntoTemplate_v2();

            // 4. Prepare data on test case excel and write into test-case (template)
            Dictionary<string, int> template_col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate: true);
            int key_col = template_col_name_list[TestCase.col_Key];
            int links_col = template_col_name_list[TestCase.col_LinkedIssue];
            int summary_col = template_col_name_list[TestCase.col_Summary];
            int status_col = template_col_name_list[TestCase.col_Status];
            int last_row = ExcelAction.Get_Range_RowNumber(ExcelAction.GetTestCaseAllRange(IsTemplate: true));
            // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
            for (int excel_row_index = TestCase.DataBeginRow; excel_row_index <= last_row; excel_row_index++)
            {
                // Make sure Key of TC contains KeyPrefix
                String key = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, key_col, IsTemplate: true);
                if (key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row

                // If Links is not empty, extend bug key into long string with font settings
                String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
                if (links != "")
                {
                    List<String> linked_issue_key_list = TestCase.Convert_LinksString_To_ListOfString(links);
                    // To remove closed issue & not-in-Jira-exported-data issue
                    // 1. prepare an empty list
                    List<String> final_id_list = new List<String>();
                    List<String> global_issue_key_list = GetGlobalIssueKey(global_issue_list);
                    // 2. Loop throught all global issues, add key of this issue into final_id_list if:
                    //     (1) key of this issue exists on linked_issue_key_list
                    //     (2) status of this issue is NOT the same as defined in "filter-status"
                    foreach (Issue issue in global_issue_list)
                    {
                        // not on the list, go the next issue
                        if (linked_issue_key_list.IndexOf(issue.Key) < 0)
                        {
                            continue;
                        }
                        // status the same as one of those defined in "filter-status", go to next issue
                        if (fileter_status_list.IndexOf(issue.Status) >= 0)
                        {
                            continue;
                        }
                        // 2 checks are passed, add into final_id_list.Add
                        final_id_list.Add(issue.Key);
                    }
                    // 
                    List<StyleString> str_list = StyleString.ExtendIssueDescription(final_id_list, global_issue_description_list);
                    // write into template excel
                    ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
                }

                // 4.x update Status according to judgement report
                // search judgement report within list generated in 2.1
                // if found, get judgement value and update to Status
                //String judgement = GetJudgementValue(workbook, worksheet);
                //ExcelAction.SetCellValue(worksheet, excel_row_index, status_col, judgement);  
                String summary = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, summary_col, IsTemplate: true);
                String[] sp_str = summary.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                String worksheet_name = sp_str[0];

                if (report_list.ContainsKey(worksheet_name))
                {
                    String current_status = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, status_col, IsTemplate: true);
                    if (current_status == TestCase.STR_FINISHED)
                    {
                        // If current_status is "Finished" in excel report, it will be updated according to judgement of corresponding test report.
                        String workbook_filename = report_list[worksheet_name];
                        String judgement_str;
                        // If judgement value is available, update it.
                        if (KeywordReport.GetJudgementValue(workbook_filename, worksheet_name, out judgement_str))
                        {
                            ExcelAction.SetTestCaseCell(excel_row_index, status_col, judgement_str, IsTemplate: true);
                        }
                    }
                    else
                    {
                        // no update at the moment
                    }
                }
            }

            // 5. auto-fit-height of column links
            ExcelAction.TestCase_AutoFit_Column(links_col, IsTemplate: true);

            // 6. Write to another filename with datetime
            string dest_filename = Storage.GenerateFilenameWithDateTime(tclist_filename, FileExt: ".xlsx");
            ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename, IsTemplate: true);

            // Close Test Case Excel
            ExcelAction.CloseTestCaseExcel();
        }

        static public ExcelAction.ExcelStatus WriteBacktoTCJiraExcel_OpenExcel(String tclist_filename, String template_filename)
        {
            ExcelAction.ExcelStatus status;

            status = ExcelAction.OpenTestCaseExcel(tclist_filename);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                ExcelAction.CloseTestCaseExcel();
                return status;
            }

            status = ExcelAction.OpenTestCaseExcel(template_filename, IsTemplate: true);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                ExcelAction.CloseTestCaseExcel();
                return status;
            }
            return status;
        }

        static public Dictionary<String, String> WriteBacktoTCJiraExcel_GetReportList(String judgement_report_dir)
        {
            Dictionary<String, String> report_list = new Dictionary<String, String>();

            if (judgement_report_dir != "")
            {
                List<String> file_list = Storage.ListFilesUnderDirectory(judgement_report_dir);
                foreach (String name in file_list)
                {
                    // File existing check protection (it is better also checked and giving warning before entering this function)
                    if (Storage.FileExists(name) == false)
                        continue; // no warning here, simply skip this file.

                    String full_filename = Storage.GetFullPath(name);
                    String sheet_name = TestPlan.GetSheetNameAccordingToFilename(name);
                    try
                    {
                        report_list.Add(sheet_name, full_filename);
                    }
                    catch (ArgumentException)
                    {
                        Console.WriteLine("Sheet name:" + sheet_name + " already exists.");
                    }

                }
            }
            return report_list;
        }

        static public Boolean WriteBacktoTCJiraExcel_GetLinkedIssueResult(String links, out List<StyleString> Link_Issue_Detail)
        {
            Boolean ret = false;
            Link_Issue_Detail = new List<StyleString>();

            if (links != "")
            {
                List<String> linked_issue_key_list = TestCase.Convert_LinksString_To_ListOfString(links);
                // To remove closed issue & not-in-Jira-exported-data issue
                // 1. prepare an empty list
                List<String> final_id_list = new List<String>();
                List<String> global_issue_key_list = GetGlobalIssueKey(global_issue_list);
                // 2. Loop throught all global issues, add key of this issue into final_id_list if:
                //     (1) key of this issue exists on linked_issue_key_list
                //     (2) status of this issue is NOT the same as defined in "filter-status"
                foreach (Issue issue in global_issue_list)
                {
                    // not on the list, go the next issue
                    if (linked_issue_key_list.IndexOf(issue.Key) < 0)
                    {
                        continue;
                    }
                    // status the same as one of those defined in "filter-status", go to next issue
                    if (fileter_status_list.IndexOf(issue.Status) >= 0)
                    {
                        continue;
                    }
                    // 2 checks are passed, add into final_id_list.Add
                    final_id_list.Add(issue.Key);
                }
                // 
                Link_Issue_Detail = StyleString.ExtendIssueDescription(final_id_list, global_issue_description_list);
                ret = true;
            }
            return ret;
        }

        static public Boolean WriteBacktoTCJiraExcel_NeedStatusUpdateValueAccordingToJudgement
                (String status, String report_name, Dictionary<String, String> report_list, out String judgement_string)
        {
            Boolean ret = false;
            
            // if report is available, get judgement string
            String summary = report_name;
            String worksheet_name = TestPlan.GetSheetNameAccordingToSummary(summary);
            judgement_string = ""; // empty if cannot get judgement value
            if (report_list.ContainsKey(worksheet_name))
            {
                // If current_status is "Finished" in excel report, it will be updated according to judgement of corresponding test report.
                String workbook_filename = report_list[worksheet_name];
                String judgement_str;
                // If judgement value is available, update it.
                if (KeywordReport.GetJudgementValue(workbook_filename, worksheet_name, out judgement_str))
                {
                    judgement_string = judgement_str;
                }
            }

            if (status == TestCase.STR_FINISHED)
            {
                ret = true;
            }
            else
            {
                ret = false;
            }

            return ret;
        }

        // Split some part of V2 into sub-functions 
        static public void WriteBacktoTCJiraExcelV3(String tclist_filename, String template_filename, String judgement_report_dir = "")
        {
            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            ExcelAction.ExcelStatus status;

            // 1. open test case excel
            status = WriteBacktoTCJiraExcel_OpenExcel(tclist_filename, template_filename);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                return; // to-be-checked if here
            }

            // 2. Get report_list under judgement_report_dir
            Dictionary<String, String> report_list = new Dictionary<String, String>();
            report_list = WriteBacktoTCJiraExcel_GetReportList(judgement_report_dir);

            // 3. Copy test case data into template excel -- both will have the same row/col and (almost) same data
            ExcelAction.CopyTestCaseIntoTemplate_v2();

            // 4. Prepare data on test case excel and write into test-case (template)
            Dictionary<string, int> template_col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate: true);
            int key_col = template_col_name_list[TestCase.col_Key];
            int links_col = template_col_name_list[TestCase.col_LinkedIssue];
            int summary_col = template_col_name_list[TestCase.col_Summary];
            int status_col = template_col_name_list[TestCase.col_Status];
            int last_row = ExcelAction.Get_Range_RowNumber(ExcelAction.GetTestCaseAllRange(IsTemplate: true));
            // for 4.3 & 4.4
            int col_end = ExcelAction.GetTestCaseExcelRange_Col(IsTemplate: true);
            List<TestPlanKeyword> keyword_list = KeywordReport.GetGlobalKeywordList();
            // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
            for (int excel_row_index = TestCase.DataBeginRow; excel_row_index <= last_row; excel_row_index++)
            {
                // Make sure Key of TC contains KeyPrefix
                String key = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, key_col, IsTemplate: true);
                if (key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row
                String report_name = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, summary_col, IsTemplate: true);
                if (report_name == "") { break; } // 2nd protection to prevent not a TC row

                // 4.1 Extend bug key string (if not empty) into long string with font settings
                String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
                if (links != "")
                {
                    List<StyleString> str_list;
                    WriteBacktoTCJiraExcel_GetLinkedIssueResult(links, out str_list);
                    ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
                }

                // 4.2 update Status (if it is Finished) according to judgement report (if report is available)
                String current_status = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, status_col, IsTemplate: true);
                String judgement_str;
                Boolean update_status = false;
                update_status = WriteBacktoTCJiraExcel_NeedStatusUpdateValueAccordingToJudgement(current_status, report_name, report_list, out judgement_str);
                if (update_status)
                {
                    // update only of judgement_string is available.
                    if (judgement_str != "")
                    {
                        ExcelAction.SetTestCaseCell(excel_row_index, status_col, judgement_str, IsTemplate: true);
                    }
                }

                // 4.3 always fill judgement value for reference outside report border.
                ExcelAction.SetTestCaseCell(excel_row_index, (col_end+1), judgement_str, IsTemplate: true);

                // 4.4 
                // get buglist from keyword report and show it.
                String worksheet_name = TestPlan.GetSheetNameAccordingToSummary(report_name);
                String workbook_name = report_name;
                // check only worksheet because original workbook full name can't be recover here
                List<TestPlanKeyword> ws_keyword_list = KeywordReport.FilterSingleReportKeyword(keyword_list, "", worksheet_name); 

                if (ws_keyword_list.Count > 0)
                {
                    List<StyleString> str_list = new List<StyleString>();
                    StyleString new_line_str = new StyleString("\n");
                    foreach (TestPlanKeyword keyword in ws_keyword_list)
                    {
                        // Only write to keyword on currently open sheet
                        //if (keyword.Worksheet == sheet_name)
                        {
                            if (keyword.IssueDescriptionList.Count > 0)
                            {
                                // write issue description list
                                str_list.AddRange(keyword.IssueDescriptionList);
                                str_list.Add(new_line_str);
                            }
                        }
                    }
                    if (str_list.Count > 0) { str_list.RemoveAt(str_list.Count - 1); } // remove last '\n'
                    ExcelAction.TestCase_WriteStyleString(excel_row_index, (col_end+2), str_list, IsTemplate: true);
                }
            }

            // 5. auto-fit-height of column links
            ExcelAction.TestCase_AutoFit_Column(links_col, IsTemplate: true);

            // 6. Write to another filename with datetime
            string dest_filename = Storage.GenerateFilenameWithDateTime(tclist_filename, FileExt: ".xlsx");
            ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename, IsTemplate: true);

            // Close Test Case Excel
            ExcelAction.CloseTestCaseExcel();
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
        static String[] CloseStatusString = { Issue.STR_CLOSE };
        static public void FindFailTCLinkedIssueAllClosed(String tclist_filename, String template_filename)
        {
            // Prepare a list of key whose status is closed (waived treated as non-closed at the moment)
            List<String> ClosedIssueKey = new List<String>();
            foreach (Issue issue in global_issue_list)
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
            List<String> tc_finished = new List<String>();                     // TC Status is Finished
            List<String> tc_testing = new List<String>();                      // TC Status is Testing
            List<String> tc_none = new List<String>();                         // TC Status is None
            List<String> tc_blocked_empty_link_issue = new List<String>();     // TC Status is Blocked AND Links are empty
            List<String> tc_blocked_some_nonclosed = new List<String>();       // TC Status is Blocked AND Links have at least one non-closed issue
            List<String> tc_blocked_all_closed = new List<String>();           // TC Status is Blocked AND Links are all closed

            // looping all TC where links are not empty
            foreach (TestCase tc in global_testcase_list) // looping
            {
                if (tc.Status == TestCase.STR_FINISHED)
                {
                    tc_finished.Add(tc.Key);
                }
                else if (tc.Status == TestCase.STR_NONE)
                {
                    tc_testing.Add(tc.Key);
                }
                else if (tc.Status == TestCase.STR_TESTING)
                {
                    tc_none.Add(tc.Key);
                }
                else if (tc.Links.Trim() == "") // fail but empty linked issue
                {
                    tc_blocked_empty_link_issue.Add(tc.Key);
                }
                else
                {
                    List<String> LinkedIssueKey = TestCase.Convert_LinksString_To_ListOfString(tc.Links);
                    IEnumerable<String> LinkIssue_CloseIssue_intersect = ClosedIssueKey.Intersect(LinkedIssueKey);
                    if (LinkIssue_CloseIssue_intersect.Count() != LinkedIssueKey.Count())
                    {
                        // One ore more linked issue are not close (or close-alike), add into this list
                        tc_blocked_some_nonclosed.Add(tc.Key);
                    }
                    else
                    {
                        tc_blocked_all_closed.Add(tc.Key);
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
                    ExcelAction.CloseTestCaseExcel();           // original test case excel is to be closed.

                    // 4. Excel processing on template excel file
                    Dictionary<string, int> template_col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate: true);
                    int DataEndRow = ExcelAction.Get_Range_RowNumber(ExcelAction.GetTestCaseAllRange(IsTemplate: true));
                    int key_col = template_col_name_list[TestCase.col_Key];

                    // Visit all rows to check if key belongs to tc_fail_all_closed
                    int hide_row_start = 0, hide_row_count = 0;
                    for (int index = TestCase.DataBeginRow; index <= DataEndRow; index++)
                    {
                        // Make sure Key of TC contains KeyPrefix
                        String key = ExcelAction.GetTestCaseCellTrimmedString(index, key_col, IsTemplate: true);
                        if (key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row

                        bool blToHide = false;
                        if (tc_blocked_all_closed.Count == 0) { blToHide = true; }
                        else if (tc_blocked_all_closed.IndexOf(key) < 0) { blToHide = true; }
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
                            ExcelAction.TestCase_Hide_Row(hide_row_start, hide_row_count, IsTemplate: true);
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
                    string dest_filename = Storage.GenerateFilenameWithDateTime(tclist_filename, FileExt: ".xlsx");
                    ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename, IsTemplate: true);
                }
            }
        }

    }
}
