using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using System.IO;

namespace ExcelReportApplication
{
    class ExcelData
    {
        // class member
        private List<String> column_name;
        private int column_name_row;
        private int data_start_row;
        private int data_end_row;
        private Worksheet worksheet;
        private Workbook workbook;
        private String csv_filename;
        private List<List<String>> data_list;
        private int status_code;
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
        public int StatusCode   // property
        {
            get { return status_code; }   // get method
            //set { status_code = value; }  // set method
        }
        public Worksheet Worksheet   // property
        {
            get { return worksheet; }   // get method
            //set { worksheet = value; }  // set method
        }
        public Workbook Workbook   // property
        {
            get { return workbook; }   // get method
            //set { workbook = value; }  // set method
        }
        public String FullFileName   // property
        {
            get { return workbook.FullName; }   // get method
            // set { workbook.FullName = value; }  // set method
        }

        // class object setup function
        private void Init()
        {
            column_name = new List<String>();
            column_name_row = data_start_row = data_end_row = status_code = 0;
            workbook = null; worksheet = null;
            data_list = new List<List<String>>();
        }
        public ExcelData() { Init(); }
        public ExcelData(Workbook workbook, String sheetname, int column_name_row, int data_start_row, int data_end_row = 0)
        {
            Init();
            column_name_row = column_name_row;
            data_start_row = data_start_row;
            data_end_row = data_end_row;
            workbook = workbook;

            if (workbook == null)
            {
                status_code = -1;
                return;
            }

            // Select and read Test Plan sheet
            Worksheet Worksheet = ExcelAction.Find_Worksheet(workbook, sheetname);
            if (Worksheet == null)
            {
                status_code = -2;
                return;
            }

            SetupColumnName();

            if (Data_Start_Row <= Column_Name_Row)
            {
                status_code = -4;
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
                while (col_index++ <= Column_Name.Count())
                {
                    String str = ExcelAction.GetCellTrimmedString(worksheet, row_index, col_index);
                    if (String.IsNullOrWhiteSpace(str) == false)
                    {
                        all_white_space = false;
                    }
                    list_str.Add(str);
                }
                if (all_white_space) break;
                data_list.Add(list_str);
            }
            status_code = 1;
        }

        // class member function
        public void SetupColumnName()
        {
            column_name = new List<String>();

            // exception prevention
            if (worksheet == null)
                return;

            if (column_name_row <= 0)
                return;

            int cell_col_index = 1;
            String str = ExcelAction.GetCellTrimmedString(worksheet, column_name_row, cell_col_index);

            while (String.IsNullOrWhiteSpace(str) == false)
            {
                if (ContainsColumn(str)) // duplicated? (already existing)
                {
                    str = "_" + str;
                }
                column_name.Add(str);
                str = ExcelAction.GetCellTrimmedString(worksheet, column_name_row, ++cell_col_index);
            }
        }
        public Boolean ContainsColumn(String name) { return (column_name.Contains(name)); }
        public int GetColumnIndex(String name) { return (column_name.IndexOf(name)); } // if not found in IndexOf, return -1 
        public String GetColumnName(int column_index) { return (OutOfColumnBoundary(column_index) ? "" : (column_name[column_index])); }
        public int ListCount() { return data_list.Count; }
        public int ColumnCount() { return column_name.Count; }
        public  Boolean OutOfListBoundary(int list_index) { return ((list_index < 0) || (list_index >= ListCount())); }
        private Boolean OutOfColumnBoundary(int column_index) { return ((column_index < 0) || (column_index >= ColumnCount())); }
        private Boolean OutOfBoundary(int list_index, int column_index) { return (OutOfListBoundary(list_index) || OutOfColumnBoundary(column_index)); }
        private String GetCell(int list_index, int column_index) { return (OutOfBoundary(list_index, column_index) ? "" : (data_list[list_index][column_index])); }
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

        
        Boolean CopyDataFrom(ExcelData src_data, List<String> src_column_list = null)
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
        public void SaveWorkbookAs(String save_filename) { ExcelAction.SaveExcelWorkbook(this.workbook, save_filename); }
        //public void CopyToNewWorkbook(String new_filename) { ExcelAction.SaveExcelWorkbook(this.workbook, save_filename); }
        public Boolean UpdateExcel()
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

}
