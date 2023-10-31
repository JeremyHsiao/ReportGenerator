using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Text.RegularExpressions;


namespace ExcelReportApplication
{
    public class ExcelData
    {
        public enum Status
        {
            INIT = 0,
            OK = 1,
            ERR_WorkSheet = -2,
            ERR_DataStartRow = -4,
        }
        // class member
        private List<String> column_name;
        private int column_name_row;
        private int data_start_row;
        private int data_end_row;
        //private Worksheet worksheet;
        //private Workbook workbook;
        private String csv_filename;
        private List<List<String>> data_list;
        private Status status_code = Status.INIT;
        public List<String> Column_Name   // property
        {
            get { return column_name; }   // get method
            set { column_name = value; }  // set method
        }
        public int Column_Name_Row   // property
        {
            get { return column_name_row; }   // get method
            set { column_name_row = value; }  // set method
        }
        public int Data_Start_Row   // property
        {
            get { return data_start_row; }   // get method
            //set { data_start_row = value; }  // set method
        }
        public int Data_End_Row   // property
        {
            get { return data_end_row; }   // get method
            //set { data_end_row = value; }  // set method
        }
        public Status StatusCode   // property
        {
            get { return status_code; }   // get method
            //set { status_code = value; }  // set method
        }
        //public Worksheet Worksheet   // property
        //{
        //    get { return worksheet; }   // get method
        //    //set { worksheet = value; }  // set method
        //}
        //public Workbook Workbook   // property
        //{
        //    get { return workbook; }   // get method
        //    //set { workbook = value; }  // set method
        //}
        //public String FullFileName   // property
        //{
        //    get { return workbook.FullName; }   // get method
        //    // set { workbook.FullName = value; }  // set method
        //}

        // class object setup function
        private void Init()
        {
            column_name = new List<String>();
            column_name_row = data_start_row = data_end_row = 0;
            status_code = Status.INIT;
//            worksheet = null;
            data_list = new List<List<String>>();
        }
        public ExcelData() { Init(); }
        public void InitFromExcel(Worksheet worksheet, int column_name_row, int data_start_row, int data_end_row = 0)
        {
            Init();
            this.column_name_row = column_name_row;
            this.data_start_row = data_start_row;
            this.data_end_row = data_end_row;

            if (worksheet == null)
            {
                status_code = Status.ERR_WorkSheet;
                return;
            }

            this.column_name = InitColumnNameFromExcel(worksheet, Column_Name_Row);

            if (Data_Start_Row <= Column_Name_Row)
            {
                status_code = Status.ERR_DataStartRow;
                return;
            }

            Boolean no_end_row = false;
            if (Data_End_Row <= 0)
                no_end_row = true;

            int row_index = Data_Start_Row;
            while ((no_end_row) || (row_index < Data_End_Row))
            {
                Boolean all_white_space = true;
                int col_index = 1;
                List<String> list_str = new List<String>();
                while (col_index <= Column_Name.Count())
                {
                    String str = ExcelAction.GetCellTrimmedString(worksheet, row_index, col_index++);
                    if (String.IsNullOrWhiteSpace(str) == false)
                    {
                        all_white_space = false;
                    }
                    list_str.Add(str);
                }
                if (all_white_space) break;
                data_list.Add(list_str);
                row_index++;
            }
            this.data_end_row = row_index - 1;
            status_code = Status.OK;
        }
        public void InitFromExcelColumns(Worksheet worksheet, List<String> column_names, int column_name_row, int data_start_row, int data_end_row = 0)
        {
            Init();
            this.column_name_row = column_name_row;
            this.data_start_row = data_start_row;
            this.data_end_row = data_end_row;

            if (worksheet == null)
            {
                status_code = Status.ERR_WorkSheet;
                return;
            }

            this.column_name.AddRange(column_names);

            if (Data_Start_Row <= Column_Name_Row)
            {
                status_code = Status.ERR_DataStartRow;
                return;
            }

            Boolean no_end_row = false;
            if (Data_End_Row <= 0)
                no_end_row = true;

            // find out col_index of column_names on input excel
            // if "-1", it means this column on excel is not used.
            List<int> col_index_list = new List<int>();
            col_index_list.Add(-1);     // column_0 is not used in excel so set to -1
            int col_index = 1;
            String str = ExcelAction.GetCellTrimmedString(worksheet, Column_Name_Row, col_index);
            while (String.IsNullOrWhiteSpace(str) == false)
            {
                int col_found = column_names.IndexOf(str);
                col_index_list.Add(col_found);
                col_index++;
                str = ExcelAction.GetCellTrimmedString(worksheet, Column_Name_Row, col_index);
            }

            int row_index = Data_Start_Row;
            while ((no_end_row) || (row_index < Data_End_Row))
            {
                Boolean all_white_space = true;
                col_index = 1;
                List<String> list_str = new List<String>();
                list_str.AddRange(new string[column_names.Count]);
                while (col_index < col_index_list.Count)  // col_index is from 1 to (col_index_list.Count-1) (because a dummy value for col_index 0 has been added)
                {
                    int member_index = col_index_list[col_index];
                    if (member_index >= 0)
                    {
                        str = ExcelAction.GetCellTrimmedString(worksheet, row_index, col_index);
                        if (String.IsNullOrWhiteSpace(str) == false)
                        {
                            all_white_space = false;
                        }
                        list_str[member_index] = str;
                    }
                    col_index++;
                }
                if (all_white_space) break;
                data_list.Add(list_str);
                row_index++;
            }
            this.data_end_row = row_index - 1;
            status_code = Status.OK;
        }

        public void WriteToExcel(Worksheet worksheet, int column_name_row, List<String> column_name_list, int data_start_row, List<List<String>> data_list)
        {
            this.column_name = column_name_list;
            this.data_list = data_list;
            WriteToExcel(worksheet, column_name_row, data_start_row);
        }

        public void WriteToExcel(Worksheet worksheet, int column_name_row, int data_start_row)
        {
            this.column_name_row = column_name_row;
            this.data_start_row = data_start_row;
            WriteToExcel(worksheet);
        }

        public void WriteToExcel(Worksheet worksheet)
        {
            int row_index = column_name_row;
            int col_index = 1;
            foreach (String name in column_name)
            {
                ExcelAction.SetCellValue(worksheet, row_index, col_index++, (String)name);
            }

            row_index = data_start_row;
            foreach (List<String> line in data_list)
            {
                col_index = 1;
                foreach (String value in line)
                {
                    ExcelAction.SetCellValue(worksheet, row_index, col_index++, (String)value);
                }
                row_index++;
            }
            data_end_row = row_index--;
        }

        // class member function
        public List<String> InitColumnNameFromExcel(Worksheet worksheet, int name_row)
        {
            column_name = new List<String>();

            // exception prevention
            if ((worksheet == null)||(name_row <= 0))
                return column_name;

            int cell_col_index = 1;
            String str = ExcelAction.GetCellTrimmedString(worksheet, name_row, cell_col_index);

            while (String.IsNullOrWhiteSpace(str) == false)
            {
                if (ContainsColumn(str)) // duplicated? (already existing)
                {
                    str = "_" + str;
                }
                column_name.Add(str);
                str = ExcelAction.GetCellTrimmedString(worksheet, name_row, ++cell_col_index);
            }

            return column_name;
        }
        public Boolean ContainsColumn(String name) { return (column_name.Contains(name)); }
        public Boolean ContainsColumns(List<String> name_list) { foreach (String name in name_list) { if (ContainsColumn(name) == false) { return false; } } return true; }
        public int GetColumnIndex(String name) { return (column_name.IndexOf(name)); } // if not found in IndexOf, return -1 
        public String GetColumnName(int column_index) { return (OutOfColumnBoundary(column_index) ? "" : (column_name[column_index])); }
        public int LineCount() { return data_list.Count; }
        public int ColumnCount() { return column_name.Count; }
        public  Boolean OutOfListBoundary(int list_index) { return ((list_index < 0) || (list_index >= LineCount())); }
        private Boolean OutOfColumnBoundary(int column_index) { return ((column_index < 0) || (column_index >= ColumnCount())); }
        private Boolean OutOfBoundary(int list_index, int column_index) { return (OutOfListBoundary(list_index) || OutOfColumnBoundary(column_index)); }
        public String GetCell(int list_index, int column_index) { return (OutOfBoundary(list_index, column_index) ? "" : (data_list[list_index][column_index])); }
        public String GetCell(int list_index, String name) { return (ContainsColumn(name)) ? (GetCell(list_index, GetColumnIndex(name))) : ""; }
        public List<String> GetLine(int line_index) { return (OutOfListBoundary(line_index) ? new List<String>() : (data_list[line_index])); }
        public List<List<String>> GetLines(List<int> line_list)
        {
            List<List<String>> ret_list = new List<List<String>>();
            if (line_list.Count > 0)
            {
                foreach (int line in line_list)
                {
                    ret_list.Add(GetLine(line));
                }
            }
            return ret_list;
        }
        public List<String> GetColumn(String name) //return a list of column data
        {
            List<String> ret_list = new List<String>();
            if (ContainsColumn(name) == false)
            {
                int col_index = GetColumnIndex(name);
                foreach (List<String> line_data in data_list)
                {
                    ret_list.Add(line_data[col_index]);
                }
            }
            return ret_list;
        }
        public List<List<String>> GetColumns(List<String> name_list) // return multiple lists and each list contain data of single column
        {
            List<List<String>> ret_list = new List<List<String>>();
            if (name_list.Count > 0)
            {
                foreach (String name in name_list)
                {
                    ret_list.Add(GetColumn(name));
                }
            }
            return ret_list;
        }
        public Boolean InsertLine(int before_index = -1, List<String> insert_data = null)   // before_index = -1 means inserted at the end
        {
            Boolean b_Ret = true;

            // insert row data if available (but limited to current column count
            List<String> new_line_to_insert = new List<String>();
            if (insert_data != null)
            {
                new_line_to_insert.AddRange(insert_data);
            }
            int diff = new_line_to_insert.Count - this.ColumnCount();
            if (diff > 0)
            {
                new_line_to_insert.RemoveRange(this.ColumnCount(), diff);
            }
            else if (diff < 0)
            {
                new_line_to_insert.AddRange(new string[-diff]);
            }

            int insert_index = (OutOfColumnBoundary(before_index)) ? (this.ColumnCount()) : before_index;
            data_list.Insert(insert_index, new_line_to_insert);

            return b_Ret;
        }
        public Boolean InsertLines(int insert_lines, int before_index = -1, List<List<String>> insert_data = null)
        {
            Boolean b_Ret = true;
            int insert_index = (OutOfColumnBoundary(before_index)) ? (this.ColumnCount()) : before_index;
            foreach (var line in insert_data)
            {
                b_Ret &= InsertLine(insert_index++, line);
            }
            return b_Ret;
        }
        public Boolean WriteColumn(String column_name, List<String> write_data = null)
        {
            Boolean b_Ret = true;

            // return false if insert_name already exists
            if (ContainsColumn(column_name)) return false;

            int write_index = column_name.IndexOf(column_name);

            // insert column data or empty string
            int write_data_count = (write_data != null) ? write_data.Count : 0;
            int current_data_line_count = data_list.Count;

            // If data_to_insert is more than current data_list line, enlarge data_list
            if (write_data_count > current_data_line_count)
            {
                int diff = write_data_count - current_data_line_count;
                b_Ret &= InsertLines(diff);
            }

            // Start to insert column (first together with insert_data first)
            int line_index = 0;
            while (line_index < write_data_count)
            {
                data_list[line_index][write_index] = write_data[line_index];
                line_index++;
            }

            // here is for the case when insert_data is fewer than data_list
            while (line_index < data_list.Count)
            {
                data_list[line_index][write_index] = "";
                line_index++;
            }

            return b_Ret;
        }
        public Boolean WriteLine(int line, List<String> write_data = null)
        {
            Boolean b_Ret = true;
            if (OutOfListBoundary(line))
                return false;

            int write_data_count;
            if (write_data != null)
            {
                write_data_count = write_data.Count;
                data_list[line] = write_data;
            }
            else
            {
                write_data_count = 0;
                data_list[line].Clear();
            }

            int diff = write_data_count - this.ColumnCount();
            if (diff > 0)
            {
                data_list[line].RemoveRange(this.ColumnCount(), diff);
            }
            else if (diff < 0)
            {
                data_list[line].AddRange(new string[-diff]);
            }

            return b_Ret;
        }
        public Boolean DeleteLine(int line)
        {
            Boolean b_Ret = true;
            if (OutOfListBoundary(line))
                return false;

            data_list.RemoveAt(line);

            return b_Ret;
        }

        private Boolean ascending = true;
        private int Compare_index = -1;
        public int Compare_Sheetname_Ascending(List<String> line_x, List<String> line_y) 
        { 
            if (Compare_index<0)
                return 0;       // cannot compare

            String x = line_x[Compare_index], y = line_y[Compare_index];
            return TestPlan.Compare_Sheetname_Ascending(x, y);
        }
        public Boolean Setup_Compare_Field_and_Function(String field, Boolean ascending = true)
        {
            Boolean b_Ret = false;

            Compare_index = GetColumnIndex(field);
            if (Compare_index >= 0) b_Ret = true;
            this.ascending = ascending;
            return b_Ret;
        }

        private List<Comparison<String>> ascending_list = null;
        private List<int> Compare_index_list = null;
        //public int Compare_Ascending
        //public Boolean Setup_Compare_Field_and_Function

         //public Boolean InsertColumn(String insert_name, String before_name = null, List<String> insert_data = null)   // Before_column == null means inserted after all columns
        //{
        //    Boolean b_Ret = true;

        //    // return false if insert_name already exists
        //    if (ContainsColumn(insert_name)) return false;

        //    // decide insert_index by before_name
        //    int insert_index = (before_name != null) ? (column_name.IndexOf(before_name)) : (column_name.Count) ;

        //    // insert column_name 
        //    column_name.Insert(insert_index, insert_name);

        //    // insert column data or empty string
        //    int insert_data_count = (insert_data!=null)? insert_data.Count : 0;
        //    int current_data_line_count = data_list.Count;

        //    // If data_to_insert is more than current data_list line, enlarge data_list
        //    if ( insert_data_count > current_data_line_count )
        //    {
        //        int diff = insert_data_count - current_data_line_count;
        //        b_Ret &= InsertLines(diff);
        //    }

        //    // Start to insert column (first together with insert_data first)
        //    int line_index = 0;
        //    while ( line_index < insert_data_count )
        //    {
        //        data_list[line_index].Insert(insert_index,insert_data[line_index];
        //        line_index++;
        //    }

        //    // here is for the case when insert_data is fewer than data_list
        //    while (line_index < data_list.Count)
        //    {
        //        data_list[line_index].Insert(insert_index,"");
        //        line_index++;
        //    }

        //    return b_Ret;
        //}
        public Boolean CopyDataFrom(ExcelData src_data, List<String> src_column_list = null)
        {
            Boolean b_Ret = true;
            foreach (String name in src_column_list)
            {
                WriteColumn(name, src_data.GetColumn(name));
            }
            return b_Ret;
        }

        //Boolean Dummy(List<String> src_column_list, ExcelData dest_data)
        //{
        //    Boolean b_Ret = true;
        //    return b_Ret;
        //}

        // class member function for saving files
        public Boolean ReadFromCSV(String csv_filename)
        {
            Boolean b_ret = true;
            using (TextFieldParser csvParser = new TextFieldParser(csv_filename))
            {
                Init();

                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                if (csvParser.EndOfData)
                    return false;

                column_name.AddRange(csvParser.ReadFields());
                column_name_row = 1;

                if (csvParser.EndOfData)
                    return false;

                List<String> elements = new List<String>();
                elements.AddRange(csvParser.ReadFields());
                InsertLine(insert_data: elements);
                data_end_row = data_start_row = 2;

                while (!csvParser.EndOfData)
                {
                    elements.Clear();
                    elements.AddRange(csvParser.ReadFields());
                    data_end_row++;
                }
            }

            this.csv_filename = csv_filename;
            return b_ret;
        }
        public Boolean WriteToCSV(String Asfilename)
        {
            Boolean b_ret = true;
            var csv = new StringBuilder();
            csv.AppendLine(ListToCSVLine(column_name));
            foreach (List<String> list in data_list)
            {
                csv.AppendLine(ListToCSVLine(list));
            }
            File.WriteAllText(Asfilename, csv.ToString(), Encoding.UTF8);
            return b_ret;
        }
        static public String ListToCSVLine(List<String> list_of_string)
        {
            String ret_str = "";
            foreach (String str in list_of_string)
            {
                ret_str += ManPower.AddQuoteWithComma(str);
            }
            if (list_of_string.Count > 0)
            {
                ret_str.Remove(ret_str.Length - 1); // remove last ','
            }
            return ret_str;
        }
        
        // *** Use with caution and make sure the content of file is consistent with data_list
        //public void CopyToNewWorkbook(String new_filename) { ExcelAction.SaveExcelWorkbook(this.workbook, save_filename); }
        public Boolean UpdateExcel(Worksheet worksheet)
        {
            Boolean b_Ret = true;
            int row_index = data_start_row;
            foreach (var line in data_list)
            {
                int col_index = 1;
                foreach (var text in line)
                {
                    ExcelAction.SetCellValue(worksheet, row_index, col_index, text);
                    col_index++;
                }
                row_index++;
            }
            data_end_row = row_index;
            return b_Ret;
        }
    }

    //class TestPlanSheet
    //{
    //    static private ExcelDataInfo Excel_Info = new ExcelDataInfo();
    //    // Test Plan
    //    private static Boolean Do_Is_Selected(TestPlanSheet sheet)
    //    {
    //        const String DoOrNot = "Do or Not";
    //        int do_index = TestPlanSheet.Excel_Info.Column_Name.IndexOf(DoOrNot);

    //        if (do_index < 0)
    //            return false;

    //        String Do_or_Not_string = sheet.Excel_Data[do_index];
    //        return (Do_or_Not_string == DoOrNot);
    //    }
    //}

    public class ExcelDataApplication
    {

        static public string SheetName_TestPlan = "TestPlan";
        static public string SheetName_ImportToJira_Template = "ImportToJira_Template";
        static public string SheetName_ReportList = "ReportList";
        static public string SheetName_Jira_TestCase = "general_report";
        static public Worksheet ws_TestPlan;
        static public Worksheet ws_ImportToJira_Template;
        static public Worksheet ws_ReportList;
        static public Worksheet ws_Jira_TestCase;
        static public Boolean CheckTestPlanAndReportListAndImportToJira(Workbook workbook)
        {
            Boolean b_ret = ExcelAction.WorksheetExist(workbook, SheetName_TestPlan);
            if (b_ret)
            {
                ws_TestPlan = ExcelAction.Find_Worksheet(workbook, SheetName_TestPlan);
                b_ret = ExcelAction.WorksheetExist(workbook, SheetName_ReportList);
            }
            if (b_ret)
            {
                ws_ReportList = ExcelAction.Find_Worksheet(workbook, SheetName_ReportList);
                b_ret = ExcelAction.WorksheetExist(workbook, SheetName_ImportToJira_Template);
            }
            if (b_ret)
            {
                ws_ImportToJira_Template = ExcelAction.Find_Worksheet(workbook, SheetName_ImportToJira_Template);
            }

            return b_ret;
        }

        // Fill ImportToJira_Template (MUST check if columns exist in advance)
        // 1. Init ImportToJira_Template from worksheet -- to get all columns
        // 2. init TestPlan from worksheet with columns (see end of line) and remove lines where "Do or Not" is not "V"  { "Test Group", "Summary", "Customer", "SW version", "HW version", "Test Plan Ver.","Priority" };
        // 3. Write columns from (2) to ImportToJira_Template
        // 4. Get Column of Testcase {  "Summary", "Test Case Category", "Test Case Purpose", "Test Case Criteria" };
        // 5. Iterate all "Summary" on Testcase, check if value exists on ImportToJira_Template, if yes then write rest of data column on ImportToJira_Template
        // 6. Get "Source Report" and "Asignee" Column of Report List
        // 7. Iterate all "Source Report" on Report List, check if value exists on "Summary" of ImportToJira_Template, if yes then write "Assignee" column on ImportToJira_Template (shortened format of assignee)
        // 8. Writeback to CSV

        static public void ProcessImportToJira(String csv_filename)
        {
            // 1.
            ExcelData import_to_jira = new ExcelData();
            import_to_jira.InitFromExcel(worksheet: ws_ImportToJira_Template, column_name_row: 1, data_start_row: 2);

            // 2.
            ExcelData testplan_selected = new ExcelData();
            string[] column_to_copy_from_TestPlan = { "Test Group", "Summary", "Customer", "SW version", "HW version", "Test Plan Ver.","Priority" };
            testplan_selected.InitFromExcelColumns(worksheet: ws_TestPlan, column_names: column_to_copy_from_TestPlan.ToList(), column_name_row: 2, data_start_row: 3);
            List<String> do_or_not_list = testplan_selected.GetColumn("Do or Not");
            for (int line_index = do_or_not_list.Count - 1; line_index >= 0; line_index--)
            {
                if (do_or_not_list[line_index] != "V")
                {
                    testplan_selected.DeleteLine(line_index);
                }
            }

            // 3.
            List<List<String>> testplan_columns = testplan_selected.GetColumns(column_to_copy_from_TestPlan.ToList());
            for(int index = 0; index < column_to_copy_from_TestPlan.Count(); index++)
            {
                import_to_jira.WriteColumn(column_to_copy_from_TestPlan[index], testplan_columns[index]);
            }

            // 4.
            ExcelData jira_testcase = new ExcelData(); // to be replaced by actual input data
            string[] column_to_copy_from_Jira_TestCase = { "Summary", "Test Case Category", "Test Case Purpose", "Test Case Criteria" };
            List<List<String>> testcase_columns = jira_testcase.GetColumns(column_to_copy_from_Jira_TestCase.ToList());

            // 5.
            string ImportToJira_Key = column_to_copy_from_TestPlan[1]; // "Summary"
            int summary_index = import_to_jira.GetColumnIndex(ImportToJira_Key);
            string TestCase_key = column_to_copy_from_Jira_TestCase[0];
            int testcase_summary_index = column_to_copy_from_Jira_TestCase.ToList().IndexOf(TestCase_key);
            int index1 = import_to_jira.GetColumnIndex(column_to_copy_from_Jira_TestCase[1]);
            int index2 = import_to_jira.GetColumnIndex(column_to_copy_from_Jira_TestCase[2]);
            int index3 = import_to_jira.GetColumnIndex(column_to_copy_from_Jira_TestCase[3]);
            for (int line_index = 0; line_index < import_to_jira.LineCount(); line_index++)
            {
                List<String> line = import_to_jira.GetLine(line_index);
                String summary = line[summary_index];
                int testcase_line_index = testcase_columns[testcase_summary_index].IndexOf(summary);
                if (testcase_line_index >= 0)
                {
                    line[index1] = testcase_columns[1][testcase_line_index];    // here testcase_column is a line*1 column vectore
                    line[index2] = testcase_columns[2][testcase_line_index];    // here testcase_column is a line*1 column vectore
                    line[index3] = testcase_columns[3][testcase_line_index];    // here testcase_column is a line*1 column vectore
                }
                import_to_jira.WriteLine(line_index, line);
            }

            // 6. 
            ExcelData report_list = new ExcelData();
            string[] column_to_copy_from_ReportList = { "Source Report", "Assignee" };
            List<List<String>> reportlist_columns = report_list.GetColumns(column_to_copy_from_ReportList.ToList());

            // 7.
            ImportToJira_Key = column_to_copy_from_TestPlan[1]; // "Summary"
            summary_index = import_to_jira.GetColumnIndex(ImportToJira_Key);
            string ReportList_key = column_to_copy_from_ReportList[0];
            int reportlist_summary_index = column_to_copy_from_ReportList.ToList().IndexOf(ReportList_key);
            index1 = import_to_jira.GetColumnIndex(column_to_copy_from_ReportList[1]);
            for (int line_index = 0; line_index < import_to_jira.LineCount(); line_index++)
            {
                List<String> line = import_to_jira.GetLine(line_index);
                String summary = line[summary_index];
                int reportlist_line_index = reportlist_columns[reportlist_summary_index].IndexOf(summary);
                if (reportlist_line_index >= 0)
                {
                    String full_name = reportlist_columns[1][reportlist_line_index]; 
                    String english_name = Regex.Replace(full_name, "[\u4E00-\u9FFF]", ""); // 移除中文
                    String shortend_name = Regex.Replace(english_name, @"\W", "");
                    line[index1] = shortend_name;    // here testcase_column is a line*1 column vectore
                }
                import_to_jira.WriteLine(line_index, line);
            }

            // 8.
            import_to_jira.WriteToCSV(csv_filename);
        }

    }
}
