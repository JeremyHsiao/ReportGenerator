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
        static private List<Issue> global_issue_list = new List<Issue>();
        //static public Dictionary<string, List<StyleString>> global_full_issue_description_list = new Dictionary<string, List<StyleString>>();  // SaveIssueToSummaryReport
        //static public Dictionary<string, List<StyleString>> global_issue_description_list = new Dictionary<string, List<StyleString>>(); // TC-related
        //static public Dictionary<string, List<StyleString>> global_issue_description_list_severity = new Dictionary<string, List<StyleString>>(); //keyword-related
        static private List<TestCase> global_testcase_list = new List<TestCase>();
        static public List<String> fileter_status_list = new List<String>();
        static public List<ReportFileRecord> excel_not_report_log = new List<ReportFileRecord>();

        //static public Dictionary<string, Issue> lookup_BugList = new Dictionary<string, Issue>();
        static private Dictionary<string, TestCase> lookup_TestCase_by_Key = new Dictionary<string, TestCase>();
        static private Dictionary<string, TestCase> lookup_TestCase_by_Summary = new Dictionary<string, TestCase>();

        static public Dictionary<string, TestCase> GetTestcaseLUT_by_Sheetname()
        {
            return lookup_TestCase_by_Summary;
        }

        static public void SetTestcaseLUT_by_Sheetname(Dictionary<string, TestCase> new_tc_lut)
        {
            lookup_TestCase_by_Summary = new_tc_lut;
        }

        static public void ClearTestcaseLUT_by_Sheetname()
        {
            lookup_TestCase_by_Summary.Clear();
        }

        static public Dictionary<string, TestCase> GetTestcaseLUT_by_Key()
        {
            return lookup_TestCase_by_Key;
        }

        static public void SetTestcaseLUT_by_Key(Dictionary<string, TestCase> new_tc_lut)
        {
            lookup_TestCase_by_Key = new_tc_lut;
        }

        static public void ClearTestcaseLUT_by_Key()
        {
            lookup_TestCase_by_Key.Clear();
        }

        static public List<Issue> ReadGlobalIssueList()
        {
            return global_issue_list;
        }

        static public void WriteGlobalIssueList(List<Issue> new_issue_list)
        {
            global_issue_list = new_issue_list;
        }

        static public Boolean IsGlobalIssueListEmpty()
        {
            return (global_issue_list.Count <= 0);
        }

        static public void ClearGlobalIssueList()
        {
            global_issue_list.Clear();
        }

        static public List<TestCase> ReadGlobalTestcaseList()
        {
            return global_testcase_list;
        }

        static public void WriteGlobalTestcaseList(List<TestCase> new_tc_list)
        {
            global_testcase_list = new_tc_list;
        }

        static public Boolean IsGlobalTestcaseListEmpty()
        {
            return (global_testcase_list.Count <= 0);
        }

        static public void ClearGlobalTestcaseList()
        {
            global_testcase_list.Clear();
        }

        //private static List<String> GetGlobalIssueKey(List<Issue> issue_list)
        //{
        //    List<String> key_list = new List<String>();
        //    foreach (Issue issue in issue_list)
        //    {
        //        key_list.Add(issue.Key);
        //    }
        //    return key_list;
        //}

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
        /*
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
                    //List<String> global_issue_key_list = GetGlobalIssueKey(global_issue_list);
                    List<String> global_issue_key_list = lookup_BugList.Keys.ToList<String>();
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
        */

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
                List<String> all_file_list = Storage.ListFilesUnderDirectory(judgement_report_dir);
                List<String> file_list = Storage.FilterFilename(all_file_list);
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

        //static public Boolean ConvertBugID_to_BugDescription(String links, out List<StyleString> Link_Issue_Detail)
        //{
        //    Boolean ret = false;
        //    Link_Issue_Detail = new List<StyleString>();

        //    //if (links != "")
        //    if (String.IsNullOrWhiteSpace(links) == false)
        //    {
        //        List<String> linked_issue_key_list = TestCase.Convert_LinksString_To_ListOfString(links);
        //        // To remove closed issue & not-in-Jira-exported-data issue
        //        // 1. prepare an empty list
        //        List<String> final_id_list = new List<String>();
        //        //List<String> global_issue_key_list = GetGlobalIssueKey(global_issue_list);
        //        List<String> global_issue_key_list = lookup_BugList.Keys.ToList<String>();
        //        // 2. Loop throught all global issues, add key of this issue into final_id_list if:
        //        //     (1) key of this issue exists on linked_issue_key_list
        //        //     (2) status of this issue is NOT the same as defined in "filter-status"
        //        foreach (Issue issue in global_issue_list)
        //        {
        //            // status the same as one of those defined in "filter-status" (mostly closed issue), go to next issue
        //            if (fileter_status_list.IndexOf(issue.Status) >= 0)
        //            {
        //                continue;
        //            }
        //            // if bug id not on the list, go the next bug
        //            if (linked_issue_key_list.IndexOf(issue.Key) < 0)
        //            {
        //                continue;
        //            }
        //            // 2 checks are passed, add into final_id_list.Add
        //            final_id_list.Add(issue.Key);
        //        }
        //        // 
        //        Link_Issue_Detail = StyleString.ExtendIssueDescription(final_id_list, global_issue_description_list);
        //        ret = true;
        //    }
        //    return ret;
        //}

        static public String WriteBacktoTCJiraExcel_GetJudgementString(String worksheet_name, String workbook_filename)
        {
            String judgement_string = ""; // empty if cannot get judgement value
            // If current_status is "Finished" in excel report, it will be updated according to judgement of corresponding test report.
            String judgement_str;
            // If judgement value is available, update it.
            if (KeywordReport.GetJudgementValue(workbook_filename, worksheet_name, out judgement_str))
            {
                judgement_string = judgement_str;
            }

            return judgement_string;
        }

        // Split some part of V2 into sub-functions 
        static public void WriteBacktoTCJiraExcelV3(String tclist_filename, String template_filename, String bug_filename, List<Issue> bug_list,
            Dictionary<string, List<StyleString>> bug_description_list, String judgement_report_dir = "")
        {
            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            ExcelAction.ExcelStatus status;

            // 1. open test case excel
            status = WriteBacktoTCJiraExcel_OpenExcel(tclist_filename, template_filename);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                return; // to-be-checked if here
            }

            // 1.1 rename tc_template sheetname
            // 1.2 copy Jira and rename
            // 1.3 open bug and rename
            // 5.2 copy jira & tc worksheet
            //bug_filename
            //tclist_filename

            // 2. Get report_list under judgement_report_dir -- (sheetname, fullname)
            Dictionary<String, String> report_filelist_by_sheetname = new Dictionary<String, String>();
            report_filelist_by_sheetname = WriteBacktoTCJiraExcel_GetReportList(judgement_report_dir);

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
            Dictionary<String, List<TestPlanKeyword>> keyword_lut_by_Sheetname = KeywordReport.GenerateKeywordLUT_by_Sheetname(keyword_list);
            // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
            for (int excel_row_index = TestCase.DataBeginRow; excel_row_index <= last_row; excel_row_index++)
            {
                // Make sure Key of TC contains KeyPrefix
                String tc_key = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, key_col, IsTemplate: true);
                //if (tc_key.Contains(TestCase.KeyPrefix) == false) { break; } // If not a TC key in this row, go to next row
                //if (tc_key.Length < TestCase.KeyPrefix.Length) { continue; } // If not a TC key in this row, go to next row
                //if (String.Compare(tc_key, 0, TestCase.KeyPrefix, 0, TestCase.KeyPrefix.Length) != 0) { continue; } 
                String report_name = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, summary_col, IsTemplate: true);
                //if (String.IsNullOrWhiteSpace(report_name) == true) { continue; } // 2nd protection to prevent not a TC row
                if (TestCase.CheckValidTC_By_Key_Summary(tc_key, report_name) == false) { continue; }
                if (ReportGenerator.GetTestcaseLUT_by_Key().ContainsKey(tc_key) == false) { continue; }
                //if (report_name == "") { break; } // 2nd protection to prevent not a TC row
                String worksheet_name = TestPlan.GetSheetNameAccordingToSummary(report_name);

                // 4.1 Extend bug key string (if not empty) into long string with font settings
                String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
                //if (links != "")
                if (String.IsNullOrWhiteSpace(links) == false)
                {
                    List<StyleString> str_list;
                    str_list = StyleString.FilteredBugID_to_BugDescription(links, bug_list, bug_description_list);
                    ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
                }

                // check if report is availablea, if yes, use report to update judgement & list keyword issue of this report
                if (report_filelist_by_sheetname.ContainsKey(worksheet_name) == true)
                {
                    // 4.2 update Status (if it is Finished) according to judgement report (if report is available)

                    String current_status = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, status_col, IsTemplate: true);
                    String judgement_str;
                    //Boolean update_status = false;
                    String workbook_filename = report_filelist_by_sheetname[worksheet_name];
                    if (KeywordReport.CheckLookupReportJudgementResultExist())
                    {
                        judgement_str = KeywordReport.LookupReportJudgementResult(workbook_filename);
                    }
                    else
                    {
                        judgement_str = WriteBacktoTCJiraExcel_GetJudgementString(worksheet_name, workbook_filename);
                    }
                    if (current_status == TestCase.STR_FINISHED)
                    {
                        // update only of judgement_string is available.
                        //if (judgement_str != "")
                        if(String.IsNullOrWhiteSpace(judgement_str)==false)
                        {
                            ExcelAction.SetTestCaseCell(excel_row_index, status_col, judgement_str, IsTemplate: true);
                        }
                    }

                    if (KeywordReport.CheckGlobalKeywordListExist())
                    {
                        // 4.3 always fill judgement value for reference outside report border (if report is available)
                        ExcelAction.SetTestCaseCell(excel_row_index, (col_end + 1), judgement_str, IsTemplate: true);

                        // 4.4 
                        // get buglist from keyword report and show it.
                        List<TestPlanKeyword> ws_keyword_list = keyword_lut_by_Sheetname[worksheet_name];
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
                            ExcelAction.TestCase_WriteStyleString(excel_row_index, (col_end + 2), str_list, IsTemplate: true);
                        }
                    }
                }
            }

            // 5. auto-fit-height of column links
            ExcelAction.TestCase_AutoFit_Column(links_col, IsTemplate: true);

            // 6. Write to another filename with datetime (and close template file)
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
        static public void FindFailTCLinkedIssueAllClosed(String tclist_filename, String template_filename, List<Issue> bug_list)
        {
            // Prepare a list of key whose status is closed (waived treated as non-closed at the moment)
            List<String> ClosedIssueKey = new List<String>();
            foreach (Issue issue in bug_list)
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
            foreach (TestCase tc in ReadGlobalTestcaseList()) // looping
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
                else if (String.IsNullOrWhiteSpace(tc.Links)) // fail but empty linked issue
                {
                    tc_blocked_empty_link_issue.Add(tc.Key);
                }
                else
                {
                    List<String> LinkedIssueKey = Issue.Split_String_To_ListOfString(tc.Links);
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
