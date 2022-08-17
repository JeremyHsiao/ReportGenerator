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

    public class TestCase
    {
        private String key;
        private String group;
        private String summary;
        private String status;
        private String links;

        public String Key   // property
        {
            get { return key; }   // get method
            set { key = value; }  // set method
        }

        public String Group   // property
        {
            get { return group; }   // get method
            set { group = value; }  // set method
        }

        public String Summary   // property
        {
            get { return summary; }   // get method
            set { summary = value; }  // set method
        }

        public String Status   // property
        {
            get { return status; }   // get method
            set { status = value; }  // set method
        }

        public String Links   // property
        {
            get { return links; }   // get method
            set { links = value; }  // set method
        }

        public const string col_Key = "Key";
        public const string col_Group = "Test Group";
        public const string col_Summary = "Summary";
        public const string col_Status = "Status";
        public const string col_Links = "Links";

        public TestCase()
        {
        }

        public TestCase(String key, String group, String summary, String status, String links)
        {
            this.key = key; this.group = group; this.summary = summary; this.status = status; this.links = links;
        }
    }

    static class ReportWorker
    {

        // constant strings for worksheet used in this application.
        const string sheet_BUG_General_Result = "general_report";
        const string sheet_TC_Jira = "TC_BenQ27105_Result";
        const string sheet_Report_Result = "Result";

        // Key value
        const string BUG_KEY = "BENSE";
        const string TC_KEY = "TCBEN";

        static public List<TestCase> testcase_list = new List<TestCase> ();

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
                }

                ExcelAction.CloseExcelWithoutSaveChanges(myBugExcel);
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

        static public void WriteSytleString(ref Range input_range, List<StyleString> sytle_string_list)
        {
            // Fill the text into excel cell with default font settings.
            string txt_str = "";
            foreach (StyleString style_str in sytle_string_list)
            {
                txt_str += style_str.Text;
            }
            input_range.Value2 = txt_str;
            input_range.Characters.Font.Name = StyleString.default_font;
            input_range.Characters.Font.Size = StyleString.default_size;
            input_range.Characters.Font.Color = StyleString.default_color;

            // Change font settings when required for the string portion
            int chr_index = 1;
            foreach (StyleString style_str in sytle_string_list)
            {
                int len = style_str.Text.Length;
                if (style_str.FontPropertyChanged == true)
                {
                    input_range.get_Characters(chr_index, len).Font.Name = style_str.Font;
                    input_range.get_Characters(chr_index, len).Font.Color = style_str.Color;
                    input_range.get_Characters(chr_index, len).Font.Size = style_str.Size;
                }
                chr_index += len;
            }
        }

        static public List<TestCase> GenerateTestCaseList(string tclist_filename)
        {
            List<TestCase> ret_tc_list = new List<TestCase>();

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
                    const int row_tc_starting = 2;
                    for (int index = row_tc_starting; index <= rngLast.Row; index++)
                    {
                        Object cell_value2;
                        String key, group, summary, status, links;

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Key]].Value2;
                        key = (cell_value2==null)?"":cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Group]].Value2;
                        group = (cell_value2==null)?"":cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Summary]].Value2;
                        summary = (cell_value2==null)?"":cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Status]].Value2;
                        status = (cell_value2==null)?"":cell_value2.ToString();

                        cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Links]].Value2;
                        links = (cell_value2==null)?"":cell_value2.ToString();

                        ret_tc_list.Add(new TestCase(key, group, summary, status, links));
                    }
                }
                ExcelAction.CloseExcelWithoutSaveChanges(myTCExcel);
                myTCExcel = null;
            }
            return ret_tc_list;
        }

        static public void ProcessTCJiraExcel(string tclist_filename, Dictionary<string, List<StyleString>> bug_list)
        {
            testcase_list = GenerateTestCaseList(tclist_filename);

            if (testcase_list.Count > 0)
            {
                // Re-arrange test-case list into dictionary of key/links pair
                Dictionary<String, String> group_note_issue = new Dictionary<String, String>();
                foreach (TestCase tc in testcase_list)
                {
                    String key = tc.Key;
                    if (key != "")
                    {
                        group_note_issue.Add(key, tc.Links);
                    }
                }

                // Copy tc file to another file
                // 1. generate a filename
                string dest_filename, ext_str = Path.GetExtension(tclist_filename);
                if (ext_str != null)
                {
                    int file_wo_ext_len = tclist_filename.Length - ext_str.Length;
                    dest_filename = tclist_filename.Substring(0, file_wo_ext_len); 
                }
                else
                {
                    dest_filename = tclist_filename;
                }
                dest_filename += "_" + DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                // 2. Copy Excel
                File.Copy(tclist_filename, dest_filename);

                // 3. Write to the other file
                if (!File.Exists(dest_filename))
                {
                    // Open excel (read-only & corrupt-load)
                    Excel.Application myTCExcel = ExcelAction.OpenPreviousExcel(dest_filename);
                    //Excel.Application myTCExcel = OpenOridnaryExcel(dest_filename);
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
                            const int row_tc_starting = 2;
                            for (int index = row_tc_starting; index <= rngLast.Row; index++)
                            {
                                Object cell_value2;

                                // Make sure Key of TC is not empty
                                cell_value2 = WorkingSheet.Cells[index, col_name_list[TestCase.col_Key]].Value2;
                                if (cell_value2 == null) { break; }
                                String key = cell_value2.ToString();
                                if (key.Contains(TC_KEY) == false) { break; }

                                // If Links is not empty, extend bug key into long string with font settings
                                Range rng = WorkingSheet.Cells[index, col_name_list[TestCase.col_Links]];
                                cell_value2 = rng.Value2;
                                if (cell_value2 != null)
                                {
                                    List<StyleString> str_list = ExtendIssueDescription(group_note_issue[key], bug_list);

                                    WriteSytleString(ref rng, str_list);
                                }
                            }

                            ExcelAction.SaveChangesAndCloseExcel(myTCExcel, dest_filename);
                        }
                        else
                        {
                            // worksheet not found, close immediately
                            ExcelAction.CloseExcelWithoutSaveChanges(myTCExcel);
                        }
                        WorkingSheet = null;
                        myTCExcel = null;
                    }
                }
            }

        }

        //
        // To Be tested
        //
        static public void ProcessTCJiraAndSaveToReport(string tclist_filename, string report_filename, Dictionary<string, List<StyleString>> bug_list)
        {
            testcase_list = GenerateTestCaseList(tclist_filename);

            if (testcase_list.Count > 0)
            {
                // Re-arrange test-case list into dictionary of summary/links pair
                Dictionary<String, String> group_note_issue = new Dictionary<String, String>();
                foreach (TestCase tc in testcase_list)
                {
                    String key = tc.Summary;
                    if(key!="")
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
                        //const int result_row_column_naming = 5;
                        //const string col_Key = "TEST   ITEM";
                        //const string col_Links = "Links";
                        //Dictionary<string, int> result_col_name_list = CreateTableColumnIndex(result_worksheet, result_row_column_naming);
                        
                        // Get the last (row,col) of excel
                        Range rngLast = result_worksheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                        const int col_group = 1; // column "A"
                        const int col_issue = 3; // column "C"
                        const int row_result_starting = 6; // row starting from 6

                        for (int index = row_result_starting; index <= rngLast.Row; index++)
                        {
                            // find out which test_group
                            Object cell_value2 = result_worksheet.Cells[index, col_group].Value2;
                            if (cell_value2 == null) { break; }
                            // Check if empty issue-list in this test_group
                            String links = group_note_issue[cell_value2.ToString().Trim()];
                            List<StyleString> str_list = ExtendIssueDescription(links, bug_list);
                            Range rng = result_worksheet.Cells[index, col_issue];
                            WriteSytleString(ref rng, str_list);
                        }

                        // Save as another file //yyyyMMddHHmmss
                        string updated_report_filename, ext_str = Path.GetExtension(report_filename);
                        if (ext_str != null)
                        {
                            int file_wo_ext_len = tclist_filename.Length - ext_str.Length;
                            updated_report_filename = report_filename.Substring(0, file_wo_ext_len) + "_" +
                                                            DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            updated_report_filename = report_filename + "_" +
                                                            DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                        }
                        ExcelAction.SaveChangesAndCloseExcel(myReportExcel, updated_report_filename);
                    }
                    else
                    {
                        // worksheet not found, close immediately
                        ExcelAction.CloseExcelWithoutSaveChanges(myReportExcel);
                    }
                }
            }
        }
    }
}
