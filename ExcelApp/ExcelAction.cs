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
    /*  Font property
* 
Background	
Returns or sets the type of background for text used in charts. Can be one of the XlBackground constants.

Bold	
True if the font is bold.

Color	
Returns or sets the primary color of the font.

ColorIndex	
Returns or sets the color of the font.

Creator	
Returns a 32-bit integer that indicates the application in which this object was created.

FontStyle	
Returns or sets the font style.

Italic	
True if the font style is italic.

OutlineFont	
True if the font is an outline font.

Shadow	
True if the font is a shadow font or if the object has a shadow.

Size	
Returns or sets the size of the font.

Strikethrough	
True if the font is struck through with a horizontal line.

Subscript	
True if the font is formatted as subscript. False by default.

Superscript	
True if the font is formatted as superscript. False by default.

ThemeColor	
Returns or sets the theme color in the applied color scheme that is associated with the specified object. Read/write Object.

ThemeFont	
Returns or sets the theme font in the applied font scheme that is associated with the specified object. Read/write XlThemeFont.

TintAndShade	
Returns or sets a Single that lightens or darkens a color.
v
Underline	
Returns or sets the type of underline applied to the font.
*/

    public class StyleString
    {
        private String text;
        private String font_name;
        private Color font_color;
        private int font_size;
        private bool font_property_changed;

        public String Text   // property
        {
            get { return text; }   // get method
            set { text = value; }  // set method
        }
        public String Font   // property
        {
            get { return font_name; }   // get method
            set { font_name = value; font_property_changed = true; }  // set method
        }
        public Color Color   // property
        {
            get { return font_color; }   // get method
            set { font_color = value; font_property_changed = true; }  // set method
        }
        public int Size   // property
        {
            get { return font_size; }   // get method
            set { font_size = value; font_property_changed = true; }  // set method
        }
        public bool FontPropertyChanged  // property
        {
            get { return font_property_changed; }   // get method
        }

        static public string default_font = "Gill Sans MT";
        static public int default_size = 10;
        static public Color default_color = System.Drawing.Color.Black;

        public StyleString()
        {
            text = "";
            SetDefaultFontProperty();
        }

        public StyleString(string string_text)
        {
            text = string_text;
            SetDefaultFontProperty();
        }

        public StyleString(string string_text, Color string_color)
        {
            text = string_text;
            font_color = string_color;
            font_name = default_font;
            font_size = default_size;
            font_property_changed = true;
        }

        public StyleString(string string_text, Color string_color, string string_fontname, int string_fontsize)
        {
            text = string_text;
            font_color = string_color;
            font_name = string_fontname;
            font_size = string_fontsize;
            font_property_changed = true;
        }

        public void SetDefaultFontProperty()
        {
            font_color = Color.Black;
            font_name = default_font;
            font_size = default_size;
            font_property_changed = false;
        }
    }

    static class ExcelAction
    {
        // Open existing excel
        static public Excel.Application OpenOridnaryExcel(string filename)
        {
            // Open excel (read-only)
            Excel.Application myBugExcel = new Excel.Application();
            Workbook working_book = myBugExcel.Workbooks.Open(@filename, ReadOnly: true);
            myBugExcel.Visible = true;
            return myBugExcel;
        }

        static public Excel.Application OpenPreviousExcel(string filename)
        {
            // Open excel (read-only & corrupt-load)
            Excel.Application myBugExcel = new Excel.Application();
            //Workbook working_book = myBugExcel.Workbooks.Open(@filename)
            //Workbook working_book = myBugExcel.Workbooks.Open(@filename, ReadOnly: true, CorruptLoad: XlCorruptLoad.xlExtractData);
            myBugExcel.Workbooks.Open(@filename, ReadOnly: true, CorruptLoad: XlCorruptLoad.xlExtractData);
            myBugExcel.Visible = true;
            return myBugExcel;
        }

        // List all workshees within excel
        static public void ListSheets(Excel.Application curExcel)
        {
            int index = 0;

            Excel.Range rng = curExcel.get_Range("A1");

            foreach (Excel.Worksheet displayWorksheet in curExcel.Worksheets)
            {
                rng.get_Offset(index, 0).Value2 = displayWorksheet.Name;
                index++;
            }
        }

        // return worksheet with specified sheet_name; return null if not found
        static public Worksheet Find_Worksheet(Excel.Application curExcel, string sheet_name)
        {
            Worksheet ret = null;

            foreach (Excel.Worksheet displayWorksheet in curExcel.Worksheets)
            {
                if (displayWorksheet.Name.CompareTo(sheet_name) == 0)
                {
                    ret = displayWorksheet;
                    break;
                }
            }
            return ret;
        }
    }

    static class ReportWorker
    {

        // constant strings for worksheet used in this application.
        const string sheet_BUG_General_Result = "general_report";
        const string sheet_TC_Jira = "TC_BenQ27105_Result";

        // Key value
        const string BUG_KEY = "BENSE";
        const string TC_KEY = "TCBEN";

        static public Dictionary<string, int> CreateTableColumnIndex(Worksheet bug_worksheet, int naming_row)
        {
            // Get the last (row,col) of excel
            Range rngLast = bug_worksheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            Dictionary<string, int> col_name_list = new Dictionary<string, int>();

            for (int col_index = 1; col_index <= rngLast.Column; col_index++)
            {
                Object cell_value2 = bug_worksheet.Cells[naming_row, col_index].Value2;
                if (cell_value2 == null) { continue; }
                col_name_list.Add(cell_value2.ToString(), col_index);
            }

            return col_name_list;
        }

        static public Dictionary<string, List<StyleString>> CreateBugList(Worksheet bug_worksheet)
        {
            const string col_Key = "Key";
            const string col_Summary = "Summary";
            const string col_Severity = "Severity";
            const string col_RD_Comment = "Steps To Reproduce"; // To be updated 
            // const string col_RD_Comment = "Additional Information"; // To be updated
            //const int max_issue_no = 100000;
            Dictionary<string, List<StyleString>> bug_list = new Dictionary<string, List<StyleString>>();

            // Obtain column name listed on row 4 & its column index
            const int row_column_naming = 4;
            Dictionary<string, int> col_name_list = CreateTableColumnIndex(bug_worksheet, row_column_naming);

            // Get the last (row,col) of excel
            Range rngLast = bug_worksheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

            // Collect bug info from row 5
            const int row_starting_key_index = 5;
            for (int row_index = row_starting_key_index; row_index <= rngLast.Row; row_index++)
            {
                List<StyleString> add_style_str = new List<StyleString>();
                string add_str, key_str, summary_str, severity_str, rd_commment_str;
                Object cell_value2;

                // Get Key string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_Key]].Value2;
                if (cell_value2 == null) { continue; }
                key_str = cell_value2.ToString();
                if (key_str.Contains(BUG_KEY) == false) { continue; }

                // Get Summary string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_Summary]].Value2;
                if (cell_value2 == null) { summary_str = ""; }
                else
                {
                    summary_str = cell_value2.ToString();
                    if (!summary_str.Substring(summary_str.Length - 1, 1).Equals(".")) { summary_str += "."; }  // make sure '.' at the end
                }

                // Get Severity string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_Severity]].Value2;
                if (cell_value2 == null) { severity_str = ""; }
                else { severity_str = cell_value2.ToString(); }

                // Get Description string
                cell_value2 = bug_worksheet.Cells[row_index, col_name_list[col_RD_Comment]].Value2;
                if (cell_value2 == null) { rd_commment_str = ""; }
                else
                {
                    rd_commment_str = cell_value2.ToString();
                    // Remove strings after 1st '\n'
                    int new_line_pos = rd_commment_str.IndexOf('\n');
                    if (new_line_pos > 0)
                    {
                        rd_commment_str = rd_commment_str.Substring(0, new_line_pos);
                    }
                }

                // setup id/string pair
                StyleString str;

                add_str = key_str + summary_str + "(" + severity_str + ")";
                str = new StyleString(add_str, Color.Red);
                add_style_str.Add(str);
                if (rd_commment_str.CompareTo("") != 0)
                {
                    add_str = " --> " + rd_commment_str;
                    str = new StyleString(add_str, Color.Blue);
                    add_style_str.Add(str);
                }
                bug_list.Add(key_str, add_style_str);
            }
            return bug_list;
        }

        static public Dictionary<string, List<StyleString>> ProcessJiraBugFile(string buglist_filename)
        {
            Dictionary<string, List<StyleString>> myBug_list = new Dictionary<string, List<StyleString>>();

            // Open excel (read-only & corrupt-load)
            Excel.Application myBugExcel = ExcelAction.OpenPreviousExcel(buglist_filename);
            if (myBugExcel != null)
            {
                // Find bug worksheet and generate list of bug description string
                Worksheet WorkingSheet = ExcelAction.Find_Worksheet(myBugExcel, sheet_BUG_General_Result);
                if (WorkingSheet != null)
                {
                    myBug_list = CreateBugList(WorkingSheet);
                    myBugExcel.ActiveWorkbook.Close(SaveChanges: false);
                }

                myBugExcel.Quit();
                //釋放Excel資源 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myBugExcel);
                WorkingSheet = null;
                myBugExcel = null;
            }
            return myBug_list;
        }

        static public List<StyleString> ExtendIssueDescription(string links_str, Dictionary<string, List<StyleString>> bug_list)
        {
            List<StyleString> extended_str = new List<StyleString>();

            // protection
            if ((links_str == null) || (bug_list == null)) return null;

            // Separate keys
            string[] separators = { "," };
            string[] issues = links_str.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            // replace key with full description and combine into one string
            foreach (string key in issues)
            {
                string trimmed_key = key.Trim();
                StyleString new_line_str = new StyleString("\n");
                if (bug_list.ContainsKey(trimmed_key))
                {
                    List<StyleString> bug_str = bug_list[trimmed_key]; /// why 46 has count 2?

                    foreach (StyleString style_str in bug_str)
                    {

                        extended_str.Add(style_str);
                    }
                }
                else
                {
                    StyleString def_str = new StyleString(trimmed_key);
                    extended_str.Add(def_str);
                }
                extended_str.Add(new_line_str);
            }
            extended_str.RemoveAt(extended_str.Count - 1);  // remove last '\n'

            return extended_str;
        }

        static public void CreateTestReport()
        {
        }

        static public void ProcessTCJiraExcel(string tclist_filename, Dictionary<string, List<StyleString>> bug_list)
        {
            // Open excel (read-only & corrupt-load)
            Excel.Application myTCExcel = ExcelAction.OpenPreviousExcel(tclist_filename);
            //Excel.Application myTCExcel = OpenOridnaryExcel(tclist_filename);
            if (myTCExcel != null)
            {
                Worksheet WorkingSheet = ExcelAction.Find_Worksheet(myTCExcel, sheet_TC_Jira);
                if (WorkingSheet != null)
                {
                    const int row_column_naming = 1;
                    Dictionary<string, int> col_name_list = CreateTableColumnIndex(WorkingSheet, row_column_naming);

                    // Get the last (row,col) of excel
                    Range rngLast = WorkingSheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                    // Visit all rows and replace Bug-ID with long description of Bug.
                    const string col_Key = "Key";
                    const string col_Links = "Links";
                    const int row_tc_starting = 2;
                    for (int index = row_tc_starting; index <= rngLast.Row; index++)
                    {
                        Object cell_value2;

                        // Make sure Key of TC is not empty
                        cell_value2 = WorkingSheet.Cells[index, col_name_list[col_Key]].Value2;
                        if (cell_value2 == null) { break; }
                        if (cell_value2.ToString().Contains(TC_KEY) == false) { break; }

                        // If Links is not empty, extend bug key into long string with font settings
                        Range rng = WorkingSheet.Cells[index, col_name_list[col_Links]];
                        cell_value2 = rng.Value2;
                        if (cell_value2 != null)
                        {
                            List<StyleString> str_list = ExtendIssueDescription(cell_value2.ToString(), bug_list);

                            // Fill the text into excel cell with default font settings.
                            string txt_str = "";
                            foreach (StyleString style_str in str_list)
                            {
                                txt_str += style_str.Text;
                            }
                            rng.Value2 = txt_str;
                            rng.Characters.Font.Name = StyleString.default_font;
                            rng.Characters.Font.Size = StyleString.default_size;
                            rng.Characters.Font.Color = StyleString.default_color;

                            // Change font settings when required for the string portion
                            int chr_index = 1;
                            foreach (StyleString style_str in str_list)
                            {
                                int len = style_str.Text.Length;
                                if (style_str.FontPropertyChanged == true)
                                {
                                    rng.get_Characters(chr_index, len).Font.Name = style_str.Font;
                                    rng.get_Characters(chr_index, len).Font.Color = style_str.Color;
                                    rng.get_Characters(chr_index, len).Font.Size = style_str.Size;
                                }
                                chr_index += len;
                            }
                        }
                    }

                    // Save as another file //yyyyMMddHHmmss
                    string updated_tc_list_filename, ext_str = Path.GetExtension(tclist_filename);
                    if (ext_str != null)
                    {
                        int file_wo_ext_len = tclist_filename.Length - ext_str.Length;
                        updated_tc_list_filename = tclist_filename.Substring(0, file_wo_ext_len) + "_" +
                                                            DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                        myTCExcel.ActiveWorkbook.Close(SaveChanges: true, Filename: updated_tc_list_filename);
                    }

                }

                myTCExcel.Quit();
                //釋放Excel資源 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myTCExcel);
                WorkingSheet = null;
                myTCExcel = null;
            }
        }
    }

}
