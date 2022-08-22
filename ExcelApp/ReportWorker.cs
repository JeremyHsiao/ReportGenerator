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
        private FontStyle font_style;
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
        public FontStyle FontStyle   // property
        {
            get { return font_style; }   // get method
            set { font_style = value; font_property_changed = true; }  // set method
        }
        public bool FontPropertyChanged  // property
        {
            get { return font_property_changed; }   // get method
        }

        static public string default_font = "Gill Sans MT";
        static public int default_size = 10;
        static public Color default_color = System.Drawing.Color.Black;
        static public FontStyle default_fontstyle = FontStyle.Regular;

        public void SetProperty(Color string_color, string string_fontname, int string_fontsize, FontStyle string_fontstyle)
        {
            font_color = string_color;
            font_name = string_fontname;
            font_size = string_fontsize;
            font_style = string_fontstyle;
            font_property_changed = true;
        }

        public void SetDefaultProperty()
        {
            SetProperty(default_color, default_font, default_size, default_fontstyle);
            font_property_changed = false;
        }

        public void SetDefaultProperty(String string_text)
        {
            SetDefaultProperty();
            Text = string_text;
        }

        public StyleString()
        {
            SetDefaultProperty("");
        }

        public StyleString(string string_text)
        {
            SetDefaultProperty(string_text);
        }

        public StyleString(string string_text, Color string_color)
        {
            SetDefaultProperty(string_text);
            Color = string_color;
            text = string_text;
        }

        public StyleString(string string_text, Color string_color, string string_fontname, int string_fontsize)
        {
            SetProperty(string_color, string_fontname, string_fontsize, default_fontstyle);
            text = string_text;
        }
    }

    static class ReportWorker
    {
        static public Dictionary<string, List<StyleString>> global_bug_list = new Dictionary<string, List<StyleString>>();
        static public List<TestCase> global_testcase_list = new List<TestCase> ();

        static public Dictionary<string, List<StyleString>> ProcessBugList(string buglist_filename)
        {
            Dictionary<string, List<StyleString>> myBug_list = new Dictionary<string, List<StyleString>>();

            // Open excel (read-only & corrupt-load)
            Excel.Application myBugExcel = ExcelAction.OpenPreviousExcel(buglist_filename);
            if (myBugExcel != null)
            {
                // Find bug worksheet and generate list of bug description string
                Worksheet WorkingSheet = ExcelAction.Find_Worksheet(myBugExcel, IssueList.sheet_BUG_General_Result);
                if (WorkingSheet != null)
                {
                    myBug_list = IssueList.CreateBugListFromBugJiraFile(WorkingSheet);
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
                    List<StyleString> bug_str = bug_list[trimmed_key]; 

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
            if (extended_str.Count > 0) { extended_str.RemoveAt(extended_str.Count - 1); } // remove last '\n'
 
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
            input_range.Characters.Font.FontStyle = StyleString.default_fontstyle;

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
                    input_range.get_Characters(chr_index, len).Font.FontStyle = style_str.FontStyle;
                }
                chr_index += len;
            }
        }

        static public void SaveToReportTemplate(string report_filename)
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

            //Excel.Application myReportExcel = ExcelAction.OpenPreviousExcel(report_filename);
            Excel.Application myReportExcel = ExcelAction.OpenOridnaryExcel(report_filename);
            if (myReportExcel != null)
            {
                Worksheet result_worksheet = ExcelAction.Find_Worksheet(myReportExcel, IssueList.sheet_Report_Result);
                if (result_worksheet != null)
                {
                    //const int result_NameDefinitionRow = 5;
                    //const string col_Key = "TEST   ITEM";
                    //const string col_Links = "Links";
                    //Dictionary<string, int> result_col_name_list = CreateTableColumnIndex(result_worksheet, result_NameDefinitionRow);

                    // Get the last (row,col) of excel
                    Range rngLast = result_worksheet.get_Range("A1").SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);

                    const int col_group = 1, col_result = 2, col_issue = 3; // column "A" - "C"
                    const int row_result_starting = 6; // row starting from 6

                    for (int index = row_result_starting; index <= rngLast.Row; index++)
                    {
                        Range rng;
                        Object cell_value2; 
                        List<StyleString> str_list = new List<StyleString>();
                        String key, note;

                        // find out which test_group
                        rng = result_worksheet.Cells[index, col_group];
                        cell_value2 = rng.Value2;
                        if (cell_value2 == null) { break; } // if no value in test_group-->end of report
                        key = cell_value2.ToString();

                        // goes to next row if Result is N/A
                        rng = result_worksheet.Cells[index, col_result];
                        if (rng.Value2.ToString().Trim() == "N/A") { continue; } // goes to next row if N/A
 
                        // Get data to be filled into Note
                        // if key does not exist, Note will be empty string
                        if (!group_note_issue.TryGetValue(key, out note))
                        {
                            note = "";
                        }

                        if (note!="")
                        {
                            rng = result_worksheet.Cells[index, col_result];
                            rng.Value2 = "Fail";
                            // Fill "Note" 
                            str_list = ExtendIssueDescription(note, global_bug_list);
                            rng = result_worksheet.Cells[index, col_issue];
                            WriteSytleString(ref rng, str_list);
                        }
                        else
                        {
                            // no issue --> pass
                            rng = result_worksheet.Cells[index, col_result];
                            rng.Value2 = "Pass";
                            rng = result_worksheet.Cells[index, col_issue];
                            rng.Value2 = "";
                        }

                        // auto-fit-height of current row
                        rng.Rows.AutoFit();
                     }

                    // Save as another file with yyyyMMddHHmmss
                    string dest_filename = FileFunction.GenerateFilenameWithDateTime(report_filename);
                    ExcelAction.SaveChangesAndCloseExcel(myReportExcel, dest_filename);
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
