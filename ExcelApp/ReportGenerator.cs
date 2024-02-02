﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;


namespace ExcelReportApplication
{
    static class ReportGenerator
    {
        static private List<Issue> global_issue_list = new List<Issue>();
        //static public Dictionary<string, List<StyleString>> global_full_issue_description_list = new Dictionary<string, List<StyleString>>();  // SaveIssueToSummaryReport
        //static public Dictionary<string, List<StyleString>> global_issue_description_list = new Dictionary<string, List<StyleString>>(); // TC-related
        //static public Dictionary<string, List<StyleString>> global_issue_description_list_severity = new Dictionary<string, List<StyleString>>(); //keyword-related
        static private List<TestCase> global_testcase_list = new List<TestCase>();
        static public List<String> List_of_status_to_filter_for_tc_linked_issue = new List<String>();
        static public List<ReportFileRecord> excel_not_report_log = new List<ReportFileRecord>();

        //static public Dictionary<string, Issue> lookup_BugList = new Dictionary<string, Issue>();
        static private Dictionary<string, TestCase> lookup_TestCase_by_Key = new Dictionary<string, TestCase>();
        static private Dictionary<string, TestCase> lookup_TestCase_by_Summary = new Dictionary<string, TestCase>();

        public static String PASS_str = "Pass";
        public static String CONDITIONAL_PASS_str = "Conditional Pass";
        public static String FAIL_str = "Fail";
        public static String WAIVED_str = "Waived";

        public static String TestReport_Default_Judgement = "N/A";
        public static String TestReport_Default_Conclusion = " ";
        public static String TestReport_SaveReportByStatus = PASS_str + ", " + CONDITIONAL_PASS_str + ", " + FAIL_str; //"Pass, Conditional Pass, Fail";
        public static Boolean TestReport_ExtraSavePassReport = false;

        static public string LinkIssue_report_Font = "Gill Sans MT";
        static public int LinkIssue_report_FontSize = 12;
        static public Color LinkIssue_report_FontColor = System.Drawing.Color.Black;
        static public FontStyle LinkIssue_report_FontStyle = FontStyle.Regular;
        static public Color LinkIssue_A_Issue_Color = Color.Red;
        static public Color LinkIssue_B_Issue_Color = Color.Black;
        static public Color LinkIssue_C_Issue_Color = Color.Black;
        static public Color LinkIssue_D_Issue_Color = Color.Black;
        static public Color LinkIssue_WAIVED_ISSUE_COLOR = Color.Black;
        static public Color LinkIssue_CLOSED_ISSUE_COLOR = Color.Black;

        static public Boolean copy_bug_list = true;
        static public Boolean copy_and_extend_bug_list = false;
        static public Boolean update_status_even_no_report = false;

        public static String GetSheetNameAccordingToFilename(String filename)
        {
            String full_filename = Storage.GetFullPath(filename);
            String short_filename = Storage.GetFileName(full_filename);
            String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            String sheet_name = sp_str[0];
            return sheet_name;
        }
        public static String GetSheetNameAccordingToSummary(String summary)
        {
            String[] sp_str = summary.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            String sheet_name = sp_str[0];
            return sheet_name;
        }
        public static String GetReportTitleAccordingToFilename(String filename)
        {
            String full_filename = Storage.GetFullPath(filename);
            String short_filename_no_extension = Storage.GetFileNameWithoutExtension(full_filename);
            return short_filename_no_extension;
        }
        public static String GetReportTitleWithoutNumberAccordingToFilename(String filename)
        {
            String full_filename = Storage.GetFullPath(filename);
            String short_filename_no_extension = Storage.GetFileNameWithoutExtension(full_filename);
            String[] sp_str = short_filename_no_extension.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            String ret_string = short_filename_no_extension.Substring(sp_str[0].Length + 1);

            return ret_string;
        }
        static public int Compare_Sheetname_Ascending(String sheetname_x, String sheetname_y)
        {
            int final_compare = 0;

            // process when one of sheetname is null
            if (sheetname_x == null)
            {
                if (sheetname_y == null)
                {
                    final_compare = 0;
                    return final_compare;
                }
                else
                {
                    final_compare = -1;
                    return final_compare;
                }
            }
            else if (sheetname_y == null)
            {
                final_compare = -1;
                return final_compare;
            }

            String[] subs_x = sheetname_x.Split('.');
            String[] subs_y = sheetname_y.Split('.');

            int compare_index = 0;

            while (true)
            {
                int x_value = 0, y_value = 0;

                Boolean x_no_more_point = (compare_index < subs_x.Count()) ? false : true;
                Boolean y_no_more_point = (compare_index < subs_y.Count()) ? false : true;

                // Comparison 1: reaching end of sheetname?
                if (x_no_more_point)
                {
                    if (y_no_more_point)
                    {
                        final_compare = 0;
                        break;
                    }
                    else
                    {
                        final_compare = -1;
                        break;
                    }
                }
                else if (y_no_more_point)
                {
                    final_compare = 1;
                    break;
                }

                String x_str = subs_x[compare_index];
                String y_str = subs_y[compare_index];

                Boolean x_is_value = Int32.TryParse(x_str, out x_value);
                Boolean y_is_value = Int32.TryParse(y_str, out y_value);

                // Comparison 2: comparing text vs value? default: value < text
                if (x_is_value == false)
                {
                    if (y_is_value == false)
                    {
                        final_compare = String.Compare(x_str, y_str);
                        if (final_compare != 0)
                        {
                            // break to return final_compare
                            break;
                        }
                        else
                        {
                            // maybe there are more points to compare 
                            compare_index++;
                        }
                    }
                    else
                    {
                        final_compare = 1;
                        break;
                    }
                }
                else if (y_is_value == false)
                {
                    final_compare = -1;
                    break;
                }
                else
                {

                    if (x_value < y_value)
                    {
                        final_compare = -1;
                        break;
                    }
                    else if (x_value > y_value)
                    {
                        final_compare = 1;
                        break;
                    }
                    else
                    {
                        // maybe there are more points to compare 
                        compare_index++; // go to compare next level
                    }
                }
            }
            return final_compare;
        }
        static public int Compare_Sheetname_by_Filename_Ascending(String filename_x, String filename_y)
        {

            String sheetname_x = GetSheetNameAccordingToFilename(filename_x);
            String sheetname_y = GetSheetNameAccordingToFilename(filename_y);

            return Compare_Sheetname_Ascending(sheetname_x, sheetname_y);
        }
        static public int Compare_Sheetname_by_Filename_Descending(String filename_x, String filename_y)
        {
            int compare_result_asceding = Compare_Sheetname_by_Filename_Ascending(filename_x, filename_y);
            return -compare_result_asceding;
        }

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

        static public void UpdateGlobalIssueList(List<Issue> new_issue_list)
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

        static public List<TestCase> ReadGlobalTestCaseList()
        {
            return global_testcase_list;
        }

        static public void UpdateGlobalTestcaseList(List<TestCase> new_tc_list)
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

        static private String[] separators = { "," };

        static public List<String> Split_String_To_ListOfString(String links)
        {
            List<String> ret_list = new List<String>();
            // protection
            //if ((links == null) || (links == "")) return ret_list;   // return empty new object
            if (String.IsNullOrWhiteSpace(links)) return ret_list;   // return empty new object
            // Separate keys into string[]
            String[] issues = links.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            if (issues == null) return ret_list;
            // string[] to List<String> (trimmed) and return
            foreach (String str in issues)
            {
                ret_list.Add(str.Trim());
            }
            return ret_list;
        }

        static public String Combine_ListOfString_to_String(List<String> list)
        {
            String ret = "";
            // protection
            if (list == null) return ret;
            if (list.Count == 0) return ret;
            foreach (String str in list)
            {
                ret += str + separators[0] + " ";
            }
            ret.Trim(); // remove " " at beginning & end
            if (ret[ret.Length - 1] == ',') { ret.Remove(ret.Length - 1); }// remove last "," 
            return ret;
        }

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
                    String sheet_name = ReportGenerator.GetSheetNameAccordingToFilename(name);
                    try
                    {
                        report_list.Add(sheet_name, full_filename);
                    }
                    catch (ArgumentException)
                    {
                        LogMessage.WriteLine("Sheet name:" + sheet_name + " already exists.");
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

        static public String ExtractDate(String input_string)
        {
            String ret_string = "";
            String pattern = @"(?<year>\d{4})\/(?<month>\d{2})\/(?<day>\d{2})";
            Regex rgx = new Regex(pattern);
            Match match = rgx.Match(input_string);
            if (match.Success)
            {
                ret_string = match.Groups["year"].Value;
                ret_string += match.Groups["month"].Value;
                ret_string += match.Groups["day"].Value;
            }
            return ret_string;
        }

        static public Dictionary<String, String> GenerateReportListFullnameLUTbySheetname(List<String> report_list)
        {
            Dictionary<String, String> report_list_lut = new Dictionary<String, String>();

            List<String> file_list = Storage.FilterFilename(report_list); // protection: remove non-report files (according to filename rule)
            foreach (String name in file_list)
            {
                // File existing check protection (it is better also checked and giving warning before entering this function)
                if (Storage.FileExists(name) == false)
                    continue; // no warning here, simply skip this file.

                String full_filename = Storage.GetFullPath(name);
                String sheet_name = GetSheetNameAccordingToFilename(name);
                try
                {
                    report_list_lut.Add(sheet_name, full_filename);
                }
                catch (ArgumentException)
                {
                    LogMessage.WriteLine("Sheet name:" + sheet_name + " already exists in GenerateReportListFullnameLUTbySheetname(List<String>).");
                }

            }
            return report_list_lut;
        }

        /*
        static public void WriteBacktoTCJiraExcelV3(String tclist_filename, String template_filename, String buglist_file, String judgement_report_dir = "")
        {
            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            ExcelAction.ExcelStatus status;

            // 1. open test case excel
            status = WriteBacktoTCJiraExcel_OpenExcel(tclist_filename, template_filename, buglist_file);
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
            report_filelist_by_sheetname = GenerateReportListFullnameLUTbySheetname(judgement_report_dir);

            // 3. Copy test case data into template excel -- both will have the same row/col and (almost) same data
            ExcelAction.CopyTestCaseIntoTemplate_v2();

            // 4. Prepare data on test case excel and write into test-case (template)
            Dictionary<string, int> template_col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate: true);
            int key_col = template_col_name_list[TestCase.col_Key];
            int links_col = template_col_name_list[TestCase.col_LinkedIssue];
            int summary_col = template_col_name_list[TestCase.col_Summary];
            int status_col = template_col_name_list[TestCase.col_Status];
            int purpose_col, criteria_col;
            if (template_col_name_list.TryGetValue(TestCase.col_Purpose, out purpose_col) == false)
            {
                purpose_col = 0;
            }
            if (template_col_name_list.TryGetValue(TestCase.col_Criteria, out criteria_col) == false)
            {
                criteria_col = 0;
            }
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
                String worksheet_name = ReportGenerator.GetSheetNameAccordingToSummary(report_name);

                // 4.1 Extend bug key string (if not empty) into long string with font settings
                String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
                //if (links != "")
                if (String.IsNullOrWhiteSpace(links) == false)
                {
                    List<Issue> linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                    // List of Issue filtered by status
                    List<Issue> filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.filter_status_list_linked_issue);
                    // Sort issue by Severity and Key valie
                    List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
                    // Convert list of sorted linked issue to description list
                    List<StyleString> str_list = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);
                    ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
                }

                // check if report is availablea, if yes, use report to update judgement & list keyword issue of this report
                if (report_filelist_by_sheetname.ContainsKey(worksheet_name) == true)
                {
                    // 4.2 update Status (if it is Finished) according to judgement report (if report is available)

                    String current_status = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, status_col, IsTemplate: true);
                    String judgement_str, purpose_str, criteria_str;
                    //Boolean update_status = false;
                    String workbook_filename = report_filelist_by_sheetname[worksheet_name];
                    //if (KeywordReport.CheckLookupReportJudgementResultExist()) 
                    if (KeywordReport.CheckLookupReportInformationExist()) //
                    {
                        //judgement_str = KeywordReport.LookupReportJudgementResult(workbook_filename);
                        List<String> report_info = KeywordReport.LookupReportInformation(workbook_filename);
                        judgement_str = KeywordReport.GetJudgement(report_info);
                        purpose_str = KeywordReport.GetPurpose(report_info);
                        criteria_str = KeywordReport.GetCriteria(report_info);
                    }
                    else
                    {
                        KeywordReport.GetJudgementPurposeCriteriaValue(workbook_filename, worksheet_name, out judgement_str, out purpose_str, out criteria_str);
                        //judgement_str = WriteBacktoTCJiraExcel_GetJudgementString(worksheet_name, workbook_filename);
                    }

                    // Update focus to current status cell
                    ExcelAction.TestCase_CellActivate(excel_row_index, status_col, IsTemplate: true);

                    if (current_status == TestCase.STR_FINISHED)
                    {
                        // update only of judgement_string is available.
                        //if (judgement_str != "")
                        if (String.IsNullOrWhiteSpace(judgement_str) == false)
                        {
                            ExcelAction.SetTestCaseCell(excel_row_index, status_col, judgement_str, IsTemplate: true);
                        }
                    }
                    // 4.2.1 -- update purpose and criteria
                    // check if purpose/criteria field exists and strings are not empty
                    if ((purpose_col > 0) && (String.IsNullOrWhiteSpace(purpose_str) == false))
                    {
                        ExcelAction.SetTestCaseCell(excel_row_index, purpose_col, purpose_str, IsTemplate: true);
                    }
                    if ((criteria_col > 0) && (String.IsNullOrWhiteSpace(criteria_str) == false))
                    {
                        ExcelAction.SetTestCaseCell(excel_row_index, criteria_col, criteria_str, IsTemplate: true);
                    }

                    // If keyword is available, add 2 extra columns of keyword result judgement and keyword issue list for reference
                    if (KeywordReport.CheckGlobalKeywordListExist())
                    {
                        // 4.3 always fill judgement value for reference outside report border (if report is available)
                        ExcelAction.SetTestCaseCell(excel_row_index, (col_end + 1), judgement_str, IsTemplate: true);

                        // 4.4 
                        // get buglist from keyword report and show it.

                        // but if worksheetname is not in LUT, go fornext worksheet
                        if (keyword_lut_by_Sheetname.ContainsKey(worksheet_name) == false)
                        {
                            continue;
                        }

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
        */

        private static Boolean UpdateStatusCellByJudgement_TCTemplate(int row, int col, String judgement)
        {
            Boolean b_Ret = false;

            String current_status = ExcelAction.GetTestCaseCellTrimmedString(row, col, IsTemplate: true);

            // Update focus to current status cell
            ExcelAction.TestCase_CellActivate(row, col, IsTemplate: true);

            if (current_status == TestCase.STR_FINISHED)
            {
                // update only of judgement_string is available.
                //if (judgement_str != "")
                if (String.IsNullOrWhiteSpace(judgement) == false)
                {
                    ExcelAction.SetTestCaseCell(row, col, judgement, IsTemplate: true);
                }
            }

            b_Ret = true;
            return b_Ret;
        }

        private static Boolean UpdateStatusCellByLinkedIssue_TCTemplate(int row, int col, List<Issue> linked_issue_list)
        {
            Boolean b_Ret = false;

            String current_status = ExcelAction.GetTestCaseCellTrimmedString(row, col, IsTemplate: true);

            // Update focus to current status cell
            ExcelAction.TestCase_CellActivate(row, col, IsTemplate: true);

            // Update Status to judgement result if Status is "Finished"
            if (current_status == TestCase.STR_FINISHED)
            {
                String status_string;
                status_string = TestReport.Judgement_Decision_by_Linked_Issue_List(linked_issue_list);
                ExcelAction.SetTestCaseCell(row, col, status_string, IsTemplate: true);
            }

            b_Ret = true;
            return b_Ret;
        }

        //
        // Input: report list
        // Output: (1) if report_list is empty or report_list contains no report file --> 
        //                                                                  (1) when update_status_without_report==true
        //                                                                          all test-case status are checked and updated when update_status_without_report==true
        //                                                                  (2) when update_status_without_report==false
        //                                                                          status remains unchanged
        //         (2) if report_list contains at least one report file --> (1) when update_status_without_report==true
        //                                                                          with report updated by report, without report udpated by linked issue    
        //                                                                  (2) when update_status_without_report==false
        //                                                                          only test-cases with corresponding report are checked and updated
        // 
        static public Boolean UpdateLinkedIssueStatusOnTCTemplate(List<String> report_list, Boolean update_status_without_report = false)
        {
            Boolean bRet = false;
            int key_col = 0;
            int links_col = 0;
            int summary_col = 0;
            int status_col = 0;

            // 2. Get report_list under judgement_report_dir -- (sheetname, fullname)
            Boolean report_is_available = false;
            Dictionary<String, String> report_filelist_by_sheetname = new Dictionary<String, String>();
            if (report_list != null)
            {
                if (report_list.Count > 0)
                {
                    report_filelist_by_sheetname = GenerateReportListFullnameLUTbySheetname(report_list);
                    report_is_available = (report_filelist_by_sheetname.Count > 0) ? true : false;
                }
            }
            // if no report, criteria/purpose won't affected AND all status updated according to linked issue condition

            // 4. Prepare data on test case excel and write into test-case (template)
            Dictionary<string, int> template_col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate: true);
            key_col = (template_col_name_list.ContainsKey(TestCase.col_Key)) ? template_col_name_list[TestCase.col_Key] : 0;
            links_col = (template_col_name_list.ContainsKey(TestCase.col_LinkedIssue)) ? template_col_name_list[TestCase.col_LinkedIssue] : 0;
            summary_col = (template_col_name_list.ContainsKey(TestCase.col_Summary)) ? template_col_name_list[TestCase.col_Summary] : 0;
            status_col = (template_col_name_list.ContainsKey(TestCase.col_Status)) ? template_col_name_list[TestCase.col_Status] : 0;

            int last_row = ExcelAction.Get_Range_RowNumber(ExcelAction.GetTestCaseAllRange(IsTemplate: true));
            // for 4.3 & 4.4
            int col_end = ExcelAction.GetTestCaseExcelRange_Col(IsTemplate: true);

            // Only if reports are available
            int purpose_col = 0, criteria_col = 0;
            List<ReportKeyword> keyword_list = new List<ReportKeyword>();
            Dictionary<String, List<ReportKeyword>> keyword_lut_by_Sheetname = new Dictionary<String, List<ReportKeyword>>();
            if (report_is_available)
            {
                // For filling purpose/criteria according to reports
                if (template_col_name_list.TryGetValue(TestCase.col_Purpose, out purpose_col) == false)
                {
                    purpose_col = 0;
                }
                if (template_col_name_list.TryGetValue(TestCase.col_Criteria, out criteria_col) == false)
                {
                    criteria_col = 0;
                }
                // END

                keyword_list = TestReport.GetGlobalKeywordList();
                keyword_lut_by_Sheetname = TestReport.GenerateKeywordLUT_by_Sheetname(keyword_list);
            }

            // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
            for (int excel_row_index = TestCase.DataBeginRow; excel_row_index <= last_row; excel_row_index++)
            {
                // Make sure Key of TC contains KeyPrefix (if key column is available)
                String tc_key = "";
                if (key_col > 0)
                {
                    tc_key = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, key_col, IsTemplate: true);
                    if (TestCase.CheckValidTC_By_KeyPrefix(tc_key) == false) { continue; }
                    if (ReportGenerator.GetTestcaseLUT_by_Key().ContainsKey(tc_key) == false) { continue; }
                }

                String report_name = "";
                String worksheet_name = "";
                if ((report_is_available) && (summary_col > 0))
                {
                    report_name = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, summary_col, IsTemplate: true);
                    if (String.IsNullOrWhiteSpace(report_name) == false)
                    {
                        worksheet_name = GetSheetNameAccordingToSummary(report_name);
                        if (String.IsNullOrWhiteSpace(tc_key) == false)
                        {
                            // when key value is available, check key value
                            if (TestCase.CheckValidTC_By_Key_Summary(tc_key, report_name) == false)
                            {
                                worksheet_name = "";            // clear if it doesn't pass key check, clear worksheet_name
                            }
                        }
                    }
                }

                List<Issue> linked_issue_list = new List<Issue>();
                List<Issue> filtered_linked_issue_list = new List<Issue>();
                if (links_col > 0)
                {
                    // 4.1 Extend bug key string (if not empty) into long string with font settings
                    String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
                    if (String.IsNullOrWhiteSpace(links) == false)
                    {
                        linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                        // List of Issue filtered by status
                        filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.List_of_status_to_filter_for_tc_linked_issue);
                        // Sort issue by Severity and Key value (A first then larger key first if same severity)
                        List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
                        // Convert list of sorted linked issue to description list
                        List<StyleString> str_list = Issue.BugList_ToLinkedIssueDescription(sorted_filtered_linked_issue_list);
                        ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
                    }
                }

                if ((report_is_available) && (String.IsNullOrWhiteSpace(worksheet_name) == false))
                {
                    // check if report is availablea, if yes, use report to update criterial/purpose & status (if FINISHE, status= judgement)
                    if (report_filelist_by_sheetname.ContainsKey(worksheet_name) == true)
                    {
                        // 4.2 update Status (if it is Finished) according to judgement report (if report is available)

                        String judgement_str, purpose_str, criteria_str;
                        String workbook_filename = report_filelist_by_sheetname[worksheet_name];
                        if (TestReport.CheckLookupReportInformationExist()) //
                        {
                            //judgement_str = KeywordReport.LookupReportJudgementResult(workbook_filename);
                            List<String> report_info = TestReport.LookupReportInformation(workbook_filename);
                            judgement_str = TestReport.GetJudgement(report_info);
                            purpose_str = TestReport.GetPurpose(report_info);
                            criteria_str = TestReport.GetCriteria(report_info);
                        }
                        else
                        {
                            TestReport.GetJudgementPurposeCriteriaValue(workbook_filename, worksheet_name, out judgement_str, out purpose_str, out criteria_str);
                        }

                        // update status by judgement (if status==FINISHED)
                        if (status_col > 0)
                        {
                            UpdateStatusCellByJudgement_TCTemplate(excel_row_index, status_col, judgement_str);
                        }

                        // 4.2.1 -- update purpose and criteria
                        // check if purpose/criteria field exists and strings are not empty
                        if ((purpose_col > 0) && (String.IsNullOrWhiteSpace(purpose_str) == false))
                        {
                            ExcelAction.SetTestCaseCell(excel_row_index, purpose_col, purpose_str, IsTemplate: true);
                        }
                        if ((criteria_col > 0) && (String.IsNullOrWhiteSpace(criteria_str) == false))
                        {
                            ExcelAction.SetTestCaseCell(excel_row_index, criteria_col, criteria_str, IsTemplate: true);
                        }
                    }
                    else if ((update_status_without_report) && (status_col > 0) && (links_col > 0))
                    {
                        UpdateStatusCellByLinkedIssue_TCTemplate(excel_row_index, status_col, linked_issue_list);
                    }
                }
                // For no-report case & update status even no report
                else if ((update_status_without_report) && (status_col > 0) && (links_col > 0))
                {
                    UpdateStatusCellByLinkedIssue_TCTemplate(excel_row_index, status_col, linked_issue_list);
                }
                else
                {
                    // Do nothing
                }
            }

            // 5. auto-fit-height of column links
            if (links_col > 0)
            {
                ExcelAction.TestCase_AutoFit_Column(links_col, IsTemplate: true);
            }

            bRet = true;
            return bRet;
        }

        static public Boolean ProcessBugListToExtendTestCase(Worksheet worksheet_buglist)
        {
            Boolean bRet = false;

            Dictionary<string, int> buglist_col_name_list = ExcelAction.CreateBugListColumnIndex();
            int key_col = buglist_col_name_list[Issue.col_Key];
            int links_col = buglist_col_name_list[Issue.col_LinkedIssue];
            int last_row = ExcelAction.Get_Range_RowNumber(ExcelAction.GetIssueListAllRange());
            int col_end = ExcelAction.GetBugListExcelRange_Col();

            // Visit all rows and replace Testcase Key at Linked Issue with Testcase Summary - in reverse order
            for (int excel_row_index = last_row; excel_row_index >= TestCase.DataBeginRow; excel_row_index--)
            {
                String links = ExcelAction.GetIssueListCellTrimmedString(excel_row_index, links_col);
                List<TestCase> linked_tc_list = TestCase.KeyStringToListOfTestCase(links, ReportGenerator.ReadGlobalTestCaseList());
                linked_tc_list.Reverse();
                int current_processing_index = linked_tc_list.Count;
                foreach (TestCase tc in linked_tc_list)
                {
                    // update current row
                    //List<StyleString> str_list = StyleString.TestCaseList_To_TestCaseSummary(tc.ToList());
                    List<StyleString> str_list = tc.ToTestCaseSummary();
                    StyleString.WriteStyleString(worksheet_buglist, excel_row_index, links_col, str_list);

                    // if still more rows to insert
                    if (--current_processing_index > 0)
                    {
                        // Emulate the action of copy a selected range of rows and insert/paste below the selected range 

                        // Get the rows to copy 
                        Range copyRange = worksheet_buglist.Rows[excel_row_index + ":" + excel_row_index];
                        // Insert enough new rows to fit the rows we're copying.
                        copyRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        // The copied data will be put in the same place 
                        Range dest = worksheet_buglist.Rows[excel_row_index + ":" + excel_row_index];
                        copyRange.Copy(dest);
                    }
                }
            }

            bRet = true;
            return bRet;
        }

        //static public Boolean WriteBacktoTCJiraExcelV3_rev2(String judgement_report_dir = "")
        //{

        //    // 2. Get report_list under judgement_report_dir -- (sheetname, fullname)
        //    Dictionary<String, String> report_filelist_by_sheetname = new Dictionary<String, String>();
        //    report_filelist_by_sheetname = GenerateReportListFullnameLUTbySheetname(judgement_report_dir);

        //    // 4. Prepare data on test case excel and write into test-case (template)
        //    Dictionary<string, int> template_col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate: true);
        //    int key_col = template_col_name_list[TestCase.col_Key];
        //    int links_col = template_col_name_list[TestCase.col_LinkedIssue];
        //    int summary_col = template_col_name_list[TestCase.col_Summary];
        //    int status_col = template_col_name_list[TestCase.col_Status];

        //    // For filling purpose/criteria according to reports
        //    int purpose_col, criteria_col;
        //    if (template_col_name_list.TryGetValue(TestCase.col_Purpose, out purpose_col) == false)
        //    {
        //        purpose_col = 0;
        //    }
        //    if (template_col_name_list.TryGetValue(TestCase.col_Criteria, out criteria_col) == false)
        //    {
        //        criteria_col = 0;
        //    }
        //    // END

        //    int last_row = ExcelAction.Get_Range_RowNumber(ExcelAction.GetTestCaseAllRange(IsTemplate: true));
        //    // for 4.3 & 4.4
        //    int col_end = ExcelAction.GetTestCaseExcelRange_Col(IsTemplate: true);
        //    List<TestPlanKeyword> keyword_list = KeywordReport.GetGlobalKeywordList();
        //    Dictionary<String, List<TestPlanKeyword>> keyword_lut_by_Sheetname = KeywordReport.GenerateKeywordLUT_by_Sheetname(keyword_list);

        //    // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
        //    for (int excel_row_index = TestCase.DataBeginRow; excel_row_index <= last_row; excel_row_index++)
        //    {
        //        // Make sure Key of TC contains KeyPrefix
        //        String tc_key = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, key_col, IsTemplate: true);
        //        if (TestCase.CheckValidTC_By_KeyPrefix(tc_key) == false) { continue; }
        //        if (ReportGenerator.GetTestcaseLUT_by_Key().ContainsKey(tc_key) == false) { continue; }

        //        String report_name = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, summary_col, IsTemplate: true);
        //        //if (String.IsNullOrWhiteSpace(report_name) == true) { continue; } // 2nd protection to prevent not a TC row
        //        if (TestCase.CheckValidTC_By_Key_Summary(tc_key, report_name) == false) { continue; }
        //        String worksheet_name = ReportGenerator.GetSheetNameAccordingToSummary(report_name);

        //        // 4.1 Extend bug key string (if not empty) into long string with font settings
        //        String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
        //        List<Issue> linked_issue_list = new List<Issue>();
        //        List<Issue> filtered_linked_issue_list = new List<Issue>();
        //        if (String.IsNullOrWhiteSpace(links) == false)
        //        {
        //            linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
        //            // List of Issue filtered by status
        //            filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.filter_status_list_linked_issue);
        //            // Sort issue by Severity and Key value (A first then larger key first if same severity)
        //            List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
        //            // Convert list of sorted linked issue to description list
        //            List<StyleString> str_list = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);
        //            ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
        //        }

        //        // check if report is availablea, if yes, use report to update judgement & list keyword issue of this report
        //        if (report_filelist_by_sheetname.ContainsKey(worksheet_name) == true)
        //        {
        //            // 4.2 update Status (if it is Finished) according to judgement report (if report is available)

        //            String current_status = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, status_col, IsTemplate: true);
        //            String judgement_str, purpose_str, criteria_str;
        //            //Boolean update_status = false;
        //            String workbook_filename = report_filelist_by_sheetname[worksheet_name];
        //            //if (KeywordReport.CheckLookupReportJudgementResultExist()) 
        //            if (KeywordReport.CheckLookupReportInformationExist()) //
        //            {
        //                //judgement_str = KeywordReport.LookupReportJudgementResult(workbook_filename);
        //                List<String> report_info = KeywordReport.LookupReportInformation(workbook_filename);
        //                judgement_str = KeywordReport.GetJudgement(report_info);
        //                purpose_str = KeywordReport.GetPurpose(report_info);
        //                criteria_str = KeywordReport.GetCriteria(report_info);
        //            }
        //            else
        //            {
        //                KeywordReport.GetJudgementPurposeCriteriaValue(workbook_filename, worksheet_name, out judgement_str, out purpose_str, out criteria_str);
        //                //judgement_str = WriteBacktoTCJiraExcel_GetJudgementString(worksheet_name, workbook_filename);
        //            }

        //            // Update focus to current status cell
        //            ExcelAction.TestCase_CellActivate(excel_row_index, status_col, IsTemplate: true);

        //            if (current_status == TestCase.STR_FINISHED)
        //            {
        //                // update only of judgement_string is available.
        //                //if (judgement_str != "")
        //                if (String.IsNullOrWhiteSpace(judgement_str) == false)
        //                {
        //                    ExcelAction.SetTestCaseCell(excel_row_index, status_col, judgement_str, IsTemplate: true);
        //                }
        //            }
        //            // 4.2.1 -- update purpose and criteria
        //            // check if purpose/criteria field exists and strings are not empty
        //            if ((purpose_col > 0) && (String.IsNullOrWhiteSpace(purpose_str) == false))
        //            {
        //                ExcelAction.SetTestCaseCell(excel_row_index, purpose_col, purpose_str, IsTemplate: true);
        //            }
        //            if ((criteria_col > 0) && (String.IsNullOrWhiteSpace(criteria_str) == false))
        //            {
        //                ExcelAction.SetTestCaseCell(excel_row_index, criteria_col, criteria_str, IsTemplate: true);
        //            }

        //            // If keyword is available, add 2 extra columns of keyword result judgement and keyword issue list for reference
        //            if (KeywordReport.CheckGlobalKeywordListExist())
        //            {
        //                // 4.3 always fill judgement value for reference outside report border (if report is available)
        //                ExcelAction.SetTestCaseCell(excel_row_index, (col_end + 1), judgement_str, IsTemplate: true);

        //                // 4.4 
        //                // get buglist from keyword report and show it.

        //                // but if worksheetname is not in LUT, go fornext worksheet
        //                if (keyword_lut_by_Sheetname.ContainsKey(worksheet_name) == false)
        //                {
        //                    continue;
        //                }

        //                List<TestPlanKeyword> ws_keyword_list = keyword_lut_by_Sheetname[worksheet_name];
        //                if (ws_keyword_list.Count > 0)
        //                {
        //                    List<StyleString> str_list = new List<StyleString>();
        //                    StyleString new_line_str = new StyleString("\n");
        //                    foreach (TestPlanKeyword keyword in ws_keyword_list)
        //                    {
        //                        // Only write to keyword on currently open sheet
        //                        //if (keyword.Worksheet == sheet_name)
        //                        {
        //                            if (keyword.IssueDescriptionList.Count > 0)
        //                            {
        //                                // write issue description list
        //                                str_list.AddRange(keyword.IssueDescriptionList);
        //                                str_list.Add(new_line_str);
        //                            }
        //                        }
        //                    }
        //                    if (str_list.Count > 0) { str_list.RemoveAt(str_list.Count - 1); } // remove last '\n'
        //                    ExcelAction.TestCase_WriteStyleString(excel_row_index, (col_end + 2), str_list, IsTemplate: true);
        //                }
        //            }
        //            //END

        //        }
        //    }

        //    // 5. auto-fit-height of column links
        //    ExcelAction.TestCase_AutoFit_Column(links_col, IsTemplate: true);

        //    return true;
        //}

        //

        //static public Boolean WriteBacktoTCJiraExcelV3_simplified_branch_writing_template_by_TC()
        //{
        //    // 4. Prepare data on test case excel and write into test-case (template)
        //    Dictionary<string, int> template_col_name_list = ExcelAction.CreateTestCaseColumnIndex(IsTemplate: true);
        //    int key_col = template_col_name_list[TestCase.col_Key];
        //    int links_col = template_col_name_list[TestCase.col_LinkedIssue];
        //    int summary_col = template_col_name_list[TestCase.col_Summary];
        //    int status_col = template_col_name_list[TestCase.col_Status];
        //    int last_row = ExcelAction.Get_Range_RowNumber(ExcelAction.GetTestCaseAllRange(IsTemplate: true));
        //    // for 4.3 & 4.4
        //    int col_end = ExcelAction.GetTestCaseExcelRange_Col(IsTemplate: true);

        //    // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
        //    for (int excel_row_index = TestCase.DataBeginRow; excel_row_index <= last_row; excel_row_index++)
        //    {
        //        // Make sure Key of TC contains KeyPrefix
        //        String tc_key = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, key_col, IsTemplate: true);
        //        if (TestCase.CheckValidTC_By_KeyPrefix(tc_key) == false) { continue; }
        //        if (ReportGenerator.GetTestcaseLUT_by_Key().ContainsKey(tc_key) == false) { continue; }

        //        // 4.1 Extend bug key string (if not empty) into long string with font settings
        //        String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
        //        List<Issue> linked_issue_list = new List<Issue>();
        //        List<Issue> filtered_linked_issue_list = new List<Issue>();
        //        if (String.IsNullOrWhiteSpace(links) == false)
        //        {
        //            linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
        //            // List of Issue filtered by status
        //            filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.filter_status_list_linked_issue);
        //            // Sort issue by Severity and Key value (A first then larger key first if same severity)
        //            List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
        //            // Convert list of sorted linked issue to description list
        //            List<StyleString> str_list = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);
        //            ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
        //        }

        //        // 4.2 update Status (if it is Finished) according to linked issue count

        //        String current_status = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, status_col, IsTemplate: true);

        //        // Update focus to current status cell
        //        ExcelAction.TestCase_CellActivate(excel_row_index, status_col, IsTemplate: true);

        //        // Update Status to judgement result if Status is "Finished"
        //        if (current_status == TestCase.STR_FINISHED)
        //        {
        //            String status_string;
        //            status_string = KeywordReport.Judgement_Decision_by_Linked_Issue(linked_issue_list);
        //            ExcelAction.SetTestCaseCell(excel_row_index, status_col, status_string, IsTemplate: true);
        //        }
        //    }

        //    // 5. auto-fit-height of column links
        //    ExcelAction.TestCase_AutoFit_Column(links_col, IsTemplate: true);

        //    return true;
        //}

        // Difference with V3 -- Use linked issue status to update STATUS when it is FINISHED (instead of judgement_report as in V3)
        // Do not use keyword result

        /*
        static public void WriteBacktoTCJiraExcelV3_simplified_branch(String tclist_filename, String template_filename, String buglist_file)
        {
            // Open original excel (read-only & corrupt-load) and write to another filename when closed
            ExcelAction.ExcelStatus status;

            // 1. open test case excel
            status = WriteBacktoTCJiraExcel_OpenExcel(tclist_filename, template_filename, buglist_file);
            if (status != ExcelAction.ExcelStatus.OK)
            {
                return; // to-be-checked if here
            }

            // 3. Copy test case data into template excel -- both will have the same row but column can be in different order
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

            // Visit all rows and replace Bug-ID at Linked Issue with long description of Bug.
            for (int excel_row_index = TestCase.DataBeginRow; excel_row_index <= last_row; excel_row_index++)
            {
                // Make sure Key of TC contains KeyPrefix
                String tc_key = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, key_col, IsTemplate: true);
                if (TestCase.CheckValidTC_By_KeyPrefix(tc_key) == false) { continue; }
                if (ReportGenerator.GetTestcaseLUT_by_Key().ContainsKey(tc_key) == false) { continue; }

                // 4.1 Extend bug key string (if not empty) into long string with font settings
                String links = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, links_col, IsTemplate: true);
                List<Issue> filtered_linked_issue_list = new List<Issue> ();
                if (String.IsNullOrWhiteSpace(links) == false)
                {
                    List<Issue> linked_issue_list = Issue.KeyStringToListOfIssue(links, ReportGenerator.ReadGlobalIssueList());
                    // List of Issue filtered by status
                    filtered_linked_issue_list = Issue.FilterIssueByStatus(linked_issue_list, ReportGenerator.filter_status_list_linked_issue);
                    // Sort issue by Severity and Key value (A first then larger key first if same severity)
                    List<Issue> sorted_filtered_linked_issue_list = Issue.SortingBySeverityAndKey(filtered_linked_issue_list);
                    // Convert list of sorted linked issue to description list
                    List<StyleString> str_list = StyleString.BugList_To_LinkedIssueDescription(sorted_filtered_linked_issue_list);
                    ExcelAction.TestCase_WriteStyleString(excel_row_index, links_col, str_list, IsTemplate: true);
                }

                // 4.2 update Status (if it is Finished) according to linked issue count

                String current_status = ExcelAction.GetTestCaseCellTrimmedString(excel_row_index, status_col, IsTemplate: true);

                // Update focus to current status cell
                ExcelAction.TestCase_CellActivate(excel_row_index, status_col, IsTemplate: true);

                if (current_status == TestCase.STR_FINISHED)
                {
                    String status_string;
                    if (filtered_linked_issue_list.Count == 0)
                    {
                        status_string = PASS_str;
                    }
                    else
                    {
                        status_string = FAIL_str;
                    }
                    ExcelAction.SetTestCaseCell(excel_row_index, status_col, status_string, IsTemplate: true);
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
        */

        static public Boolean ProcessBugListExcel(String buglist_file)
        {
            String buglist_filename = Storage.GetFullPath(buglist_file);
            if (!Storage.FileExists(buglist_filename))
            {
                MainForm.SystemLogAddLine("bug file doesn't exist: " + buglist_filename);
                return false;
            }

            // open bug and process bug
            List<Issue> ret_issue_list = new List<Issue>();
            Boolean bug_open = Issue.OpenBugListExcel(buglist_filename);
            if (bug_open == false)
            {
                MainForm.SystemLogAddLine("Bug file open failed");
                return false;
            }
            ret_issue_list = Issue.GenerateIssueList_processing_data();
            UpdateGlobalIssueList(ret_issue_list);
            if (ReportGenerator.IsGlobalIssueListEmpty())
            {
                MainForm.SystemLogAddLine("Empty bug-list");
                return false;
            }
            return true;
        }

        static public Boolean ProcessTeseCaseExcel(String tc_file)
        {
            String tc_filename = Storage.GetFullPath(tc_file);
            if (!Storage.FileExists(tc_filename))
            {
                MainForm.SystemLogAddLine("TestCase file doesn't exist: " + tc_filename);
                return false;
            }
            Boolean tc_open = TestCase.OpenTestCaseExcel(tc_filename);
            if (tc_open == false)
            {
                MainForm.SystemLogAddLine("TestCase file open failed");
                return false;
            }
            List<TestCase> ret_tc_list = new List<TestCase>();
            ret_tc_list = TestCase.GenerateTestCaseList_processing_data_New();
            ReportGenerator.UpdateGlobalTestcaseList(ret_tc_list);
            ReportGenerator.SetTestcaseLUT_by_Key(TestCase.UpdateTCListLUT_by_Key(ret_tc_list));
            ReportGenerator.SetTestcaseLUT_by_Sheetname(TestCase.UpdateTCListLUT_by_Sheetname(ret_tc_list));

            if (ReportGenerator.IsGlobalTestcaseListEmpty())
            {
                MainForm.SystemLogAddLine("Empty TestCase");
                return false;
            }
            return true;
        }

        static public Boolean ProcessLinkedTestCaseOnBugListExcel()
        {
            Boolean b_ret = false;

            b_ret = true;
            return b_ret;
        }

        static public Boolean ProcessLinkedBugListOnTestCaseExcel()
        {
            Boolean b_ret = false;

            b_ret = true;
            return b_ret;
        }

        static public Boolean OpenTCTemplateAndPasteBugList(String template_file)
        {
            String template_filename = Storage.GetFullPath(template_file);
            if (!Storage.FileExists(template_filename))
            {
                MainForm.SystemLogAddLine("TC Template file doesn't exist: " + template_filename);
                return false;
            }

            Boolean tc_template_open = TestCase.OpenTCTemplateExcel(template_filename);
            if (tc_template_open == false)
            {
                MainForm.SystemLogAddLine("TC Template file open failed");
                return false;
            }

            String new_Buglist_sheetname = "BugList";
            String buglist_date_string = ExcelAction.GetIssueListCellTrimmedString(3, 1);
            String extracted_buglist_date = ExtractDate(buglist_date_string);
            if (String.IsNullOrWhiteSpace(extracted_buglist_date) == false)
            {
                new_Buglist_sheetname += "_" + extracted_buglist_date;
            }

            String newTClist_sheetname = "TCList";
            String tcglist_date_string = ExcelAction.GetTestCaseCellTrimmedString(3, 1, IsTemplate: false); // Use input TC as DATE
            String extracted_tclist_date = ExtractDate(tcglist_date_string);
            if (String.IsNullOrWhiteSpace(extracted_tclist_date) == false)
            {
                newTClist_sheetname += "_" + extracted_tclist_date;
            }

            Worksheet bug_list_worksheet = ExcelAction.GetIssueListWorksheet();
            bug_list_worksheet.Name = new_Buglist_sheetname;
            // copy-and-paste into template files.
            ExcelAction.CopyBugListSheetIntoTestCaseTemplateWorkbook();

            // set template worksheet as active worksheet.
            Worksheet tc_list_worksheet = ExcelAction.GetTestCaseWorksheet(IsTemplate: true);
            tc_list_worksheet.Select();
            tc_list_worksheet.Name = newTClist_sheetname;

            return true;
        }

        static public Boolean Process_BugList_TeseCase_TCTemplate(String tc_file, String template_file, String buglist_file)
        {
            // open bug and process bug
            if (ProcessBugListExcel(buglist_file) == false)
            {
                return false;
            }

            // open tc and process tc
            if (ProcessTeseCaseExcel(tc_file) == false)
            {
                return false;
            }

            // open template and copy bug into it
            if (OpenTCTemplateAndPasteBugList(template_file) == false)
            {
                return false;
            }

            // close bug excel
            if (Issue.CloseBugListExcel() == false)
            {
                return false;
            }

            // copy tc to template
            if (ExcelAction.CopyTestCaseIntoTemplate_v2() == false)
            {
                MainForm.SystemLogAddLine("Failed @ return of OpenProcessBugExcelTeseCaseExcelTCTemplatePasteBugCloseBugPasteTC()");
                return false;
            }
            return true;
        }

        // Report 1 relocated to here
        static public Boolean Execute_ExtendLinkIssueAndUpdateStatusWithoutReport(String tc_file)
        {
            List<String> empty_report_list = new List<String>();
            // no report has been referred at all in current report 1
            if (UpdateLinkedIssueStatusOnTCTemplate(report_list: empty_report_list, update_status_without_report: true) == false)
            {
                MainForm.SystemLogAddLine("Failed @ return of Execute_ExtendLinkIssueAndUpdateStatusWithoutReport()");
                return false;
            }

            // close tc
            ExcelAction.CloseTestCaseExcel();

            // save tempalte
            // 6. Write to another filename with datetime (and close template file)
            string dest_filename = Storage.GenerateFilenameWithDateTime(tc_file, FileExt: ".xlsx");
            ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename, IsTemplate: true);

            return true;
        }

        // Report 9 & Report K
        static public Boolean Execute_UpdateLinkedIssueStatusOnTCTemplate(String tc_file, String report_dir)
        {
            Boolean b_ret = false;

            List<String> all_file_list = new List<String>();
            if (String.IsNullOrWhiteSpace(report_dir))
            {
                all_file_list = Storage.ListFilesUnderDirectory(report_dir);
            }

            b_ret = Execute_UpdateLinkedIssueStatusOnTCTemplate(tc_file, report_list: all_file_list);
            return b_ret;
        }

        // newly created version -- report_list instead of report_dir
        // Report B & Reporl L (indirectly used by report 9/K)
        static public Boolean Execute_UpdateLinkedIssueStatusOnTCTemplate(String tc_file, List<String> report_list)
        {
            Boolean b_ret = false;

            b_ret = ReportGenerator.UpdateLinkedIssueStatusOnTCTemplate(report_list: report_list, update_status_without_report: update_status_even_no_report);
            if (b_ret == false)
            {
                MainForm.SystemLogAddLine("Failed @ return of UpdateLinkedIssueStatusOnTCTemplate()");
            }

            // close tc
            ExcelAction.CloseTestCaseExcel();

            // save tempalte
            // 6. Write to another filename with datetime (and close template file)
            string dest_filename = Storage.GenerateFilenameWithDateTime(tc_file, FileExt: ".xlsx");
            ExcelAction.SaveChangesAndCloseTestCaseExcel(dest_filename, IsTemplate: true);

            return b_ret;
        }

        static public List<String> FilterReportFileListByTCStatus(List<String> file_list)
        {
            List<String> ret_file_list = new List<String>();

            foreach (String filename in file_list)
            {
                String sheetname = GetSheetNameAccordingToFilename(filename);
                if (GetTestcaseLUT_by_Sheetname().ContainsKey(sheetname))
                {
                    TestCase tc = GetTestcaseLUT_by_Sheetname()[sheetname];
                    String report_status = tc.ReturnStatusByLinkedIssue();
                    if (TestReport_SaveReportByStatus.Contains(report_status))       // if status meets one of TestReport_SaveReportByStatus
                    {
                        ret_file_list.Add(filename);
                    }
                }
            }
            return ret_file_list;
        }

        /*
        static public bool Execute_ExtendLinkIssueAndUpdateStatusByLinkIssueFilteredCount(String tc_file, String template_file, String buglist_file)
        {
            if ((ReportGenerator.IsGlobalIssueListEmpty()) || (ReportGenerator.IsGlobalTestcaseListEmpty()) ||
                (!Storage.FileExists(tc_file)) || (!Storage.FileExists(template_file) || (!Storage.FileExists(buglist_file))))
            {
                // protection check
                // Bug/TC files must have been loaded
                return false;
            }

            ReportGenerator.WriteBacktoTCJiraExcelV3_simplified_branch(tclist_filename: tc_file, template_filename: template_file, buglist_file: buglist_file);
            return true;
        }
        */


        /*
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
                        if (TestCase.CheckValidTC_By_KeyPrefix(key) == false) { break; } // If not a TC key in this row, go to next row

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
                    // Hide all not-hidden-yet rows
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

        */

        public static List<String> SplitCommaSeparatedStringIntoList(String input_string)
        {
            List<String> ret_list = new List<String>();
            String[] csv_separators = { "," };
            if (String.IsNullOrWhiteSpace(input_string) == false)
            {
                // Separate keys into string[]
                String[] issues = input_string.Split(csv_separators, StringSplitOptions.RemoveEmptyEntries);
                if (issues != null)
                {
                    // string[] to List<String> (trimmed) and return
                    foreach (String str in issues)
                    {
                        ret_list.Add(str.Trim());
                    }
                }
            }
            return ret_list;
        }
    }
}
