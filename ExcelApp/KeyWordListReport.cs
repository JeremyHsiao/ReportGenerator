using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReportApplication
{
    class ReportFileRecord
    {
        private String path;
        private String filename;
        private String expected_sheetname;
        private Boolean excelfilenameOK;
        private Boolean openfileOK;
        private Boolean findWorksheetOK;
        private Boolean findAnyKeyword;
        private Boolean otherFailure;

        //public String Path   // property
        //{
        //    get { return path; }   // get method
        //    set { path = value; }  // set method
        //}
        //public String Filename   // property
        //{
        //    get { return filename; }   // get method
        //    set { filename = value; }  // set method
        //}
        //public Boolean FilenameOK   // property
        //{
        //    get { return filenameOK; }   // get method
        //    set { filenameOK = value; }  // set method
        //}
        //public Boolean SheetnameOK   // property
        //{
        //    get { return sheetnameOK; }   // get method
        //    set { sheetnameOK = value; }  // set method
        //}
        //public Boolean ItemOK   // property
        //{
        //    get { return itemOK; }   // get method
        //    set { itemOK = value; }  // set method
        //}
        //public Boolean CaptionOK   // property
        //{
        //    get { return captionOK; }   // get method
        //    set { captionOK = value; }  // set method
        //}
        //public Boolean OtherFailure   // property
        //{
        //    get { return otherFailure; }   // get method
        //    set { otherFailure = value; }  // set method
        //}

        //
        public ReportFileRecord() { this.path = this.filename = ""; this.excelfilenameOK = this.openfileOK = this.otherFailure = false; }
        public ReportFileRecord(String path = "", String filename = "", String expected_sheetname = "")
        {
            this.path = path;
            this.filename = filename;
            this.expected_sheetname = expected_sheetname;
            this.excelfilenameOK = false;
            this.openfileOK = false;
            this.otherFailure = false;
        }

        public String GetFullFilePath()
        {
            String path, filename, ret_str;
            path = this.path;
            filename = this.filename;
            ret_str = Storage.GetValidFullFilename(path, filename);
            return ret_str;
        }

        //public ReportFileRecord(String path, String filename, Boolean filenameOK,
        //                            Boolean sheetnameOK, Boolean itemOK, Boolean captionOK, Boolean otherFailure=false)
        //{ SetRecord(path, filename, filenameOK, sheetnameOK, itemOK, captionOK, otherFailure); }

        // only set fail flag, don't change if fail_option isn't set to true
        public void SetFlagFail(Boolean excelfilenamefail = false, Boolean openfileFail = false, Boolean findWorksheetFail = false,
                                Boolean findNoKeyword = false, Boolean otherFailure = false)
        {
            if (excelfilenamefail) { this.excelfilenameOK = false; }
            if (openfileFail) { this.openfileOK = false; }
            if (findWorksheetFail) { this.findWorksheetOK = false; }
            if (findNoKeyword) { this.findAnyKeyword = false; }
            if (otherFailure) { this.otherFailure = true; }
        }
        // only set OK flag, don't change if OK_option isn't set to true
        public void SetFlagOK(Boolean excelfilenameOK = false, Boolean openfileOK = false, Boolean findWorksheetOK = false, 
                                Boolean findAnyKeyword = false, Boolean otherAllOK = false)
        {
            if (excelfilenameOK) { this.excelfilenameOK = true; }
            if (openfileOK) { this.openfileOK = true; }
            if (findWorksheetOK) { this.findWorksheetOK = true; }
            if (findAnyKeyword) { this.findAnyKeyword = true; }
            if (otherAllOK) { this.otherFailure = false; }
        }
        public void GetFlagValue(out Boolean excelfilenameOK, out Boolean openfileOK, out Boolean findWorksheetOK, 
                                out Boolean findAnyKeyword, out Boolean otherFailure)
        {
            excelfilenameOK = this.excelfilenameOK;
            openfileOK = this.openfileOK;
            findWorksheetOK = this.findWorksheetOK;
            findAnyKeyword = this.findAnyKeyword;
            otherFailure = this.otherFailure;
        }
        public void SetFlagValue(Boolean excelfilenameOK, Boolean openfileOK, Boolean findWorksheetOK, Boolean findAnyKeyword, Boolean otherFailure = false)
        {
            this.excelfilenameOK = excelfilenameOK;
            this.openfileOK = openfileOK;
            this.findWorksheetOK = findWorksheetOK;
            this.findAnyKeyword = findAnyKeyword;
            this.otherFailure = otherFailure;
        }
        public void GetRecord(out String path, out String filename, out String expected_sheetname, out Boolean excelfilenameOK, 
                            out Boolean openfileOK, out Boolean findWorksheetOK, out Boolean findAnyKeyword, out Boolean otherFailure)
        {
            path = this.path;
            filename = this.filename;
            expected_sheetname = this.expected_sheetname;
            this.GetFlagValue(out excelfilenameOK, out openfileOK, out findWorksheetOK, out findAnyKeyword, out otherFailure);
        }

        public void SetRecord(String path, String filename, String expected_sheetname, Boolean excelfilenameOK, Boolean openfileOK, 
                            Boolean findWorksheetOK, Boolean findAnyKeyword, Boolean otherFailure = false)
        {
            this.path = path;
            this.filename = filename;
            this.expected_sheetname = expected_sheetname;
            this.SetFlagValue(excelfilenameOK, openfileOK, findWorksheetOK, findAnyKeyword, otherFailure);
        }
    }

    class KeyWordListReport
    {
        //static public string Template_Excel = "KeywordLogTemplate.xlsx";
        static public string WS_KeyWord_List = "Keyword_List";
        static public string WS_NotKeyWord_File = "Not_Keyword_File";
        static public string Output_Excel = "KeywordListLog.xlsx";

        static public int keyword_list_title_row = 1;
        static public int keyword_list_title_col_start = 1;
        private static String[] keyword_list_title = new String[] 
        {
            "Keyword",
            "Filepath",
            "Filename",
            "Worksheet",
            "Duplicated?",
       };

        private static String[] keyword_list_with_issue_title = new String[] 
        {
            "Keyword",
            "Filepath",
            "Filename",
            "Worksheet",
            "Duplicated?",
            "Bug Count",
            "Bug List",
       };

        static public int not_keyword_file_title_row = 1;
        static public int not_keyword_file_title_col_start = 1;
        private static String[] not_keyword_file_title = new String[] 
        {
            "Filepath",
            "Filename",
            "FilenameOK",
            "OpenFileOK",
            "FindWorksheetOK",
            "FindAnyKeyword",
            "AnyOtherFailure",
       };

        //static public 
        static public void OutputKeywordLog(String out_path, List<TestReportKeyword> keyword_list, 
                                            List<ReportFileRecord> not_keyword_report_list, String keyword_output_filename ="", Boolean output_keyword_issue = false)
        {
            // Open template excel and write to another filename when closed
            ExcelAction.ExcelStatus status;

            // 1. open a new workbook with 2 worksheets
            status = ExcelAction.CreateNewKeywordListExcel();
            if (status != ExcelAction.ExcelStatus.OK)
            {
                ExcelAction.CloseNewKeywordListExcel();
                return; // to-be-checked if here
            }

            // 2. output keyword_list
            List<List<Object>> output_keyword_list_table = new List<List<Object>>();
            List<Object> row_list = new List<Object> ();
            // title
            if (!output_keyword_issue)
            {
                row_list.AddRange(keyword_list_title);
            }
            else
            {
                row_list.AddRange(keyword_list_with_issue_title);
            }
            output_keyword_list_table.Add(row_list);

            // list keyword string of all duplicated keyword
            List<String> duplicate_keyword_str_list = TestReport.ListDuplicatedKeywordString(keyword_list);

            foreach (TestReportKeyword keyword_data in keyword_list)
            {
                String keyword = keyword_data.Keyword;
                String full_path = keyword_data.Workbook;
                row_list = new List<Object>();
                // Keyword
                row_list.Add(keyword);
                //"Filepath"
                row_list.Add(Storage.GetDirectoryName(full_path));
                //"Filename"
                row_list.Add(Storage.GetFileName(full_path));
                //"Worksheet"
                row_list.Add(keyword_data.Worksheet);
                //"Duplicated?"
                if (duplicate_keyword_str_list.Contains(keyword))
                {
                    // duplicated
                    row_list.Add("v");
                }
                else 
                {
                    // add this blank space so that all subsequent element is placed at correct position.
                    row_list.Add(" ");
                }

                if (output_keyword_issue)
                {
                    row_list.Add(keyword_data.KeywordIssues.Count().ToString());
                    row_list.Add(StyleString.StyleStringListToString(keyword_data.IssueDescriptionList));
                }
                output_keyword_list_table.Add(row_list);
            }
            ExcelAction.WriteTableToKeywordList(output_keyword_list_table);

            // 3. output not-keyword file linst
            // title
            List<List<Object>> output_not_report_file_table = new List<List<Object>>();
            row_list.Clear();
            row_list.AddRange(not_keyword_file_title);
            output_not_report_file_table.Add(row_list);

            // not-keyword-file
            foreach (ReportFileRecord not_keyword_report in not_keyword_report_list)
            {
                String path, filename, expected_sheetname;
                Boolean excelfilenameOK,openfileOK, findWorksheetOK, findAnyKeyword, otherFailure;

                not_keyword_report.GetRecord(out path, out filename, out expected_sheetname, out excelfilenameOK, out openfileOK, 
                                            out findWorksheetOK, out findAnyKeyword, out otherFailure);

                row_list = new List<Object>();
                //"Filepath",
                row_list.Add(path);
                //"Filename",
                row_list.Add(filename);
                //"FilenameOK"
                if (!excelfilenameOK)
                {
                    row_list.Add("X");
                    row_list.Add("-");
                    row_list.Add("-");
                    row_list.Add("-");
                }
                else
                {
                    row_list.Add(" ");
                    //"OpenFileOK",
                    if (!openfileOK)
                    {
                        row_list.Add("X");
                        row_list.Add("-");
                        row_list.Add("-");
                    }
                    else
                    {
                        row_list.Add(" ");
                        //"FindWorksheetOK",
                        if (!findWorksheetOK)
                        {
                            row_list.Add("X");
                            row_list.Add("-");
                        }
                        else
                        {
                            row_list.Add(" ");
                            //"FindAnyKeyword"
                            if (!findAnyKeyword)
                            {
                                row_list.Add("X");
                            }
                            else
                            {
                                row_list.Add(" ");
                            }
                        }
                    }
                }
                //"AnyOtherFailure"
                if(otherFailure)
                {
                    row_list.Add("X");
                }
                else
                {
                    row_list.Add(" ");
                }
                output_not_report_file_table.Add(row_list);
            }
            ExcelAction.WriteTableToNotKeywordFile(output_not_report_file_table);

            if (keyword_output_filename != "")
            {
                Output_Excel = keyword_output_filename;
            }

            // 4. Close and Save
            String output_file_full_path = Storage.GetValidFullFilename(out_path, Output_Excel);
            // if parent directory does not exist, create recursively all parents
            Storage.CreateDirectory(out_path, auto_parent_dir: true);
            ExcelAction.SaveChangesAndCloseNewKeywordListExcel(output_file_full_path);
        }

    }
}
