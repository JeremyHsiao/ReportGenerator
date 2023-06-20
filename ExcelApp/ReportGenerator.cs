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

        // Must be updated if new report type added #NewReportType
        public enum ReportType
        {
            FullIssueDescription_TC = 0,
            FullIssueDescription_Summary,
            StandardTestReportCreation,
            KeywordIssue_Report_SingleFile,
            TC_Likely_Passed,
            FindAllKeywordInReport,
            KeywordIssue_Report_Directory,
            Excel_Sheet_Name_Update_Tool,
            FullIssueDescription_TC_report_judgement,
            TC_TestReportCreation,
            TC_AutoCorrectReport_By_Filename,
            TC_AutoCorrectReport_By_ExcelList, 
        }

        public static ReportType[] ReportSelectableTable =
        {
            ReportType.FullIssueDescription_TC,
            //ReportType.FullIssueDescription_Summary,
            //ReportType.StandardTestReportCreation,
            ReportType.KeywordIssue_Report_SingleFile,
            //ReportType.TC_Likely_Passed,
            ReportType.FindAllKeywordInReport,
            ReportType.KeywordIssue_Report_Directory,
            //ReportType.Excel_Sheet_Name_Update_Tool,
            ReportType.FullIssueDescription_TC_report_judgement,
            ReportType.TC_TestReportCreation,
            ReportType.TC_AutoCorrectReport_By_Filename,
            ReportType.TC_AutoCorrectReport_By_ExcelList, 
         };

        //public static ReportType[] ReportSelectableTable =
        //{
        //    ReportType.FullIssueDescription_TC,
        //    ReportType.FullIssueDescription_Summary,
        //    ReportType.StandardTestReportCreation,
        //    ReportType.KeywordIssue_Report_SingleFile,
        //    ReportType.TC_Likely_Passed,
        //    ReportType.FindAllKeywordInReport,
        //    ReportType.KeywordIssue_Report_Directory,
        //    ReportType.Excel_Sheet_Name_Update_Tool,
        //    ReportType.FullIssueDescription_TC_report_judgement,
        //    ReportType.TC_TestReportCreation,
        //    ReportType.TC_AutoCorrectReport_By_Filename,
        //    ReportType.TC_AutoCorrectReport_By_ExcelList,
        //};

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

        // Must be updated if new report type added #NewReportType
        private static String[] ReportName = new String[] 
        {
            "1.Issue Description for TC",
            "2.Issue Description for Summary",
            "3.Standard Test Report Creator",
            "4.Keyword Issue - Single File",
            "5.TC likely Pass",
            "6.List Keywords of all detailed reports",
            "7.Keyword Issue - Directory",
            "8.Excel sheet name update tool",
            "9.TC issue/judgement",
            "A.Jira Test Report Creator",
            "B.Auto-correct report header",
            "C.Report Header Update Tool ",
       };

        // Must be updated if new report type added #NewReportType
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
                "Keyword Issue to Report - Single File", 
                "Input:",  "  Test Plan/Report with Keyword",
                "Output:", "  Test Plan/Report with keyword issue list inserted on the 1st-column next to the right-side of printable area",
            },
            new String[] 
            {
                "Test case status is Fail but its linked issues are closed", 
                "Input:",  "  Issue List + Test Case + Template (for Test case output)",
                "Output:", "  Test Case whose linked issues are closed (other TC are hidden)",
            },
            new String[] 
            {
                "Go Through all report to list down all keywords", 
                "Input:",  "  Root-directory of Report Files",
                "Output:", "  All keywords listed on output excel file",
            },
            new String[] 
            {
                "Keyword Issue to Report - Files under directory", 
                "Input:",  "  Test Plan/Reports with Keyword under user-specified directory",
                "Output:", "  Test Plan/Reports with keyword issue list inserted on the 1st-column next to the right-side of printable area",
            },
            new String[] 
            {
                "Excel Reports to be checked - Files under directory", 
                "Input:",  "  Test Plan/Reports",
                "Output:", "  Test Plan/Reports with updated sheet-name if original sheet-name doesn't follow rules",
            },
            new String[] 
            {
                "Issue Description + judgement from report for TC Report", 
                "Input:",  "  Issue List + Test Case + Template (for Test case output)",
                "Output:", "  Test Case in the format of template file with linked issue in full description",
            },
            new String[] 
            {
                "Create file structure of Test Report according to TC on the Jira Test Case file", 
                "Input:",  "  Jira Test Case File & directories of source report and of output destination",
                "Output:", "  Directory structure and report files under directories",
            },
            new String[] 
            {
                "Worksheet name & 1st row (header) will be auto-corrected.", 
                "Input:",  "  Root Directory of test reports",
                "Output:", "  Auto-corrected test reports",
            },
            new String[] 
            {
                "Worksheet name & 1st row (header) of report will be renamed and these reports are copied to corresponding folders", 
                "Input:",  "  Input excel file",
                "Output:", "  Reports copied and renamed (filename / worksheet name) according to input excel file",
            },
        };

        private static String[][] UI_Label = new String[][] 
        {
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Output Template",
                "Main Test Plan",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Output Template",
                "Main Test Plan",
           },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Output Template",
                "Main Test Plant",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Single Report",
                "Main Test Plan",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Output Template",
                "Main Test Plan",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Report Path",
                "Main Test Plan",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Output Template",
                "Main Test Plan",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Output Template",
                "Main Test Plan",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Output Template",
                "Test Report Path",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Report Output Path",
                "Report Source Path",
            },
            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Check Report Path",
                "Report Source Path",
            },

            new String[] 
            {
                "Jira Issue File", 
                "Test Case File",
                "Report Output Path",
                "Input Excel File",
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

        public static String[] GetLabelTextArray(int type_index)
        {
            return UI_Label[type_index];
        }

        public static String GetReportDescription(ReportType type)
        {
            return GetReportDescription(ReportTypeToInt(type));
        }

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
                    String short_filename = Storage.GetFileName(full_filename);
                    String[] sp_str = short_filename.Split(new Char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                    String sheet_name = sp_str[0];
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
                    if (current_status == "Finished")
                    {
                        // If current_status is "Finished" in excel report, it will be updated according to judgement of corresponding test report.
                        String workbook_filename = report_list[worksheet_name];
                        String judgement_str;
                        // If judgement value is available, update it.
                        if (GetJudgementValue(workbook_filename, worksheet_name, out judgement_str))
                        {
                            ExcelAction.SetTestCaseCell(excel_row_index, status_col, judgement_str, IsTemplate: true);
                        }
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

        // This function is used to get judgement result (only read and no update to report) of test report
        static public Boolean GetJudgementValue(String report_workbook, String report_worksheet, out String judgement_str)
        {
            Boolean b_ret = false;
            String ret_str = ""; // default returning judgetment_str;

            // 1. Open Excel and find the sheet
            // File exist check is done outside
            Workbook wb_judgement = ExcelAction.OpenExcelWorkbook(report_workbook);
            if (wb_judgement == null)
            {
                ConsoleWarning("ERR: Open workbook in GetJudgementValue: " + report_workbook);
                judgement_str = ret_str;
                b_ret = false;
            }
            else
            {
                // 2 Open worksheet
                Worksheet ws_judgement = ExcelAction.Find_Worksheet(wb_judgement, report_worksheet);
                if (ws_judgement == null)
                {
                    ConsoleWarning("ERR: Open worksheet in V4: " + report_workbook + " sheet: " + report_worksheet);
                    judgement_str = ret_str;
                    b_ret = false;
                }
                else
                {
                    // 3. Get Judgement value
                    Object obj = ExcelAction.GetCellValue(ws_judgement, TestReport.Judgement_at_row, TestReport.Judgement_at_col);
                    if (obj != null)
                    {
                        judgement_str = (String)obj;
                        b_ret = true;
                    }
                    else
                    {
                        judgement_str = ret_str;
                        b_ret = false;
                    }
                }

                // Close excel if open succeeds
                ExcelAction.CloseExcelWorkbook(wb_judgement);
            }
            return b_ret;
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
