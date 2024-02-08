using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

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
        static public int default_size = 12;
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

        public StyleString(string string_text, Color string_color, string string_fontname, int string_fontsize, FontStyle string_fontstyle)
        {
            SetProperty(string_color, string_fontname, string_fontsize, string_fontstyle);
            text = string_text;
        }

        static public List<StyleString> EmptyList()
        {
            return StringToListOfStyleString("");
        }

        static public List<StyleString> WhiteSpaceList()
        {
            return StringToListOfStyleString(" ");
        }

        static public String StyleStringToString(StyleString stylestring)
        {
            return stylestring.text;
        }

        public List<StyleString> ConvertToList()
        {
            List<StyleString> ret_list = new List<StyleString>();
            ret_list.Add(this);
            return ret_list;
        }

        static public List<StyleString> StringToListOfStyleString(String intput_string)
        {
            List<StyleString> ret_list = new List<StyleString>();
            StyleString temp = new StyleString(intput_string);
            ret_list.Add(temp);
            return ret_list;
        }

        //
        // Common function related to StyleString
        //

        static public String StyleStringListToString(List<StyleString> list_stylestring)
        {
            String ret_str = "";

            foreach (StyleString style_string in list_stylestring)
            {
                ret_str += StyleStringToString(style_string);
            }
            return ret_str;
        }

        //
        // input: bug_id separated by comma
        // output: bug descriptions (one bug each line)
        //
        //static public List<StyleString> ExtendIssueDescription(string links_str, Dictionary<string, List<StyleString>> bug_description_list)
        //{
        //    List<StyleString> extended_str = new List<StyleString>();

        //    // protection
        //    if ((links_str == null) || (bug_description_list == null)) return null;

        //    List<String> id_list = Issue.Split_String_To_ListOfString(links_str);
        //    extended_str = ExtendIssueDescription(id_list, bug_description_list);
        //    return extended_str;
        //}

        //
        // input: bug_id List
        // output: bug descriptions (one bug each line)
        //
        //static public List<StyleString> ExtendIssueDescription(List<String> bug_id, Dictionary<string, List<StyleString>> bug_description_list)
        //{
        //    List<StyleString> extended_str = new List<StyleString>();

        //    // protection
        //    if ((bug_id == null) || (bug_description_list == null)) return null;

        //    // replace each bug_id with full description seperated by newline and combine into one string
        //    StyleString new_line_str = new StyleString("\n");
        //    foreach (string key in bug_id)
        //    {
        //        string trimmed_key = key.Trim();
        //        if (bug_description_list.ContainsKey(trimmed_key))
        //        {
        //            List<StyleString> bug_str = bug_description_list[trimmed_key];

        //            foreach (StyleString style_str in bug_str)
        //            {

        //                extended_str.Add(style_str);
        //            }
        //        }
        //        else
        //        {
        //            StyleString def_str = new StyleString(trimmed_key);
        //            extended_str.Add(def_str);
        //        }
        //        extended_str.Add(new_line_str);
        //    }
        //    if (extended_str.Count > 0) { extended_str.RemoveAt(extended_str.Count - 1); } // remove last '\n'

        //    return extended_str;
        //}

        //// extexd issue description when this issue is not filtered by its status
        //static public List<StyleString> FilteredBugID_to_BugDescription(String links, List<Issue> issue_list_source,
        //                                                                    Dictionary<string, List<StyleString>> bug_description_list)
        //{
        //    List<StyleString> Link_Issue_Detail = new List<StyleString>();

        //    //if (links != "")
        //    if (String.IsNullOrWhiteSpace(links) == false)
        //    {
        //        // filtered out issues whose key is not in links string
        //        List<Issue> key_issue_list = Issue.KeyStringToListOfIssue(links, issue_list_source);
        //        // To remove closed issue
        //        List<Issue> filtered_issue_list = Issue.FilterIssueByStatus(key_issue_list, ReportGenerator.filter_status_list_linked_issue);
        //        List<String> filtered_issue_key_list = Issue.ListOfIssueToListOfIssueKey(filtered_issue_list);
        //        Link_Issue_Detail = ExtendIssueDescription(filtered_issue_key_list, bug_description_list);
        //    }
        //    return Link_Issue_Detail;
        //}

        static public void WriteStyleString(ref Range input_range, List<StyleString> style_string_list, Boolean ClearContentFirst = false)
        {
            // Fill the text into excel cell with default font settings.
            string txt_str = "";
            foreach (StyleString style_str in style_string_list)
            {
                txt_str += style_str.Text;
            }
            if (ClearContentFirst)
            {
                input_range.ClearContents();
            }
            input_range.NumberFormat = "@";
            input_range.Value2 = txt_str;

            using (System.Drawing.Font fontTester =
                        new System.Drawing.Font(StyleString.default_font,
                                                StyleString.default_size,
                                                StyleString.default_fontstyle,
                                                GraphicsUnit.Pixel))
            {
                if (fontTester.Name == StyleString.default_font)
                {
                    // Font exists
                }
                else
                {
                    // Default font doesn't exist ==> no need to change font (font as it is) at all
                    return;
                }
            }
            //if (StyleString.default_font == "NoChange") return;

            input_range.Characters.Font.Name = StyleString.default_font;
            input_range.Characters.Font.Size = StyleString.default_size;
            input_range.Characters.Font.Color = StyleString.default_color;
            input_range.Characters.Font.FontStyle = StyleString.default_fontstyle;

            // Change font settings when required for the string portion
            int chr_index = 1;
            foreach (StyleString style_str in style_string_list)
            {
                int len = style_str.Text.Length;

                // Skip font-update if "NoChange";
                //if (style_str.Font == "NoChange") continue;
                using (System.Drawing.Font fontTester =
                            new System.Drawing.Font(style_str.Font,
                                                    style_str.Size,
                                                    style_str.FontStyle,
                                                    GraphicsUnit.Pixel))
                {
                    if (fontTester.Name == style_str.Font)
                    {
                        // Font exists
                    }
                    else
                    {
                        // Font doesn't exist ==> no need to change font use as it is
                        continue;
                    }
                }

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

        static public void WriteStyleString(Worksheet ws, int row, int col, List<StyleString> style_string_list, Boolean ClearContentFirst = false)
        {
            Range input_range = ws.Cells[row, col];
            if (ClearContentFirst)
            {
                input_range.MergeArea.ClearContents();
            }
            WriteStyleString(ref input_range, style_string_list);
        }

        static public List<StyleString> BugList_To_SummaryPageFullIssueDescription(List<Issue> issuelist)
        {
            Color descrption_color_issue = Color.Red;
            Color descrption_color_comment = Color.Blue;

            List<StyleString> ret_style_string = new List<StyleString>();
            int processed_count = 0;

            foreach (Issue issue in issuelist)
            {
                List<StyleString> value_style_str = new List<StyleString>();
                String key = issue.Key, rd_comment_str = issue.Comment;
                processed_count++;

                if (String.IsNullOrWhiteSpace(key) == false)
                {
                    String str = key + issue.Summary + "(" + issue.Severity + ")";
                    StyleString style_str = new StyleString(str, descrption_color_issue);
                    value_style_str.Add(style_str);

                    // Keep portion of string before first "\n"; if no "\n", keep whole string otherwise.
                    String short_comment = "";
                    if (rd_comment_str.Contains("\n"))
                    {
                        short_comment = rd_comment_str.Substring(0, rd_comment_str.IndexOf("\n"));
                    }
                    else
                    {
                        short_comment = rd_comment_str;
                    }
                    if (String.IsNullOrWhiteSpace(short_comment) == false)
                    {
                        str = " --> " + short_comment;
                        style_str = new StyleString(str, descrption_color_comment);
                        value_style_str.Add(style_str);
                    }

                    // Add whole string into return_list
                    if (processed_count < issuelist.Count())
                    {
                        ret_style_string.Add(new StyleString("\n"));
                    }
                }
            }
            return ret_style_string;
        }

        // create key/rich-text-issue-description pair.
        // 
        // Format: KEY+SUMMARY+(+SEVERITY+)
        //
        // For example: BENSE27105-99[OSD]Menu scenario-Color Gamut value incorrect Without Metadata when Sub screen(B)
        //

        //static public Dictionary<string, List<StyleString>> GenerateIssueDescription(List<Issue> issuelist)
        //{
        //    Dictionary<string, List<StyleString>> ret_list = new Dictionary<string, List<StyleString>>();

        //    foreach (Issue issue in issuelist)
        //    {
        //        List<StyleString> value_style_str = new List<StyleString>();
        //        String key = issue.Key, rd_comment_str = issue.Comment;

        //        if (key != "")
        //        {
        //            Boolean is_waived = false;
        //            if (issue.Status == Issue.STR_WAIVE)
        //            {
        //                is_waived = true;
        //            }

        //            String str = key + issue.Summary + "(" + issue.Severity + ")";
        //            if (is_waived)
        //            {
        //                str += "(" + KeywordReport.WAIVED_str + ")";
        //            }
        //            StyleString style_str = new StyleString(str, ReportGenerator.LinkIssue_report_FontColor);
        //            value_style_str.Add(style_str);
        //            /*
        //            // Keep portion of string before first "\n"; if no "\n", keep whole string otherwise.
        //            String short_comment = "";
        //            if (rd_comment_str.Contains("\n"))
        //            {
        //                short_comment = rd_comment_str.Substring(0, rd_comment_str.IndexOf("\n"));
        //            }
        //            else
        //            {
        //                short_comment = rd_comment_str;
        //            }
        //            if (short_comment != "")
        //            {
        //                str = " --> " + short_comment;
        //                style_str = new StyleString(str, descrption_color_comment);
        //                value_style_str.Add(style_str);
        //            }
        //            */
        //            // Add whole string into return_list
        //            ret_list.Add(key, value_style_str);
        //        }
        //    }
        //    return ret_list;
        //}

        static public List<StyleString> BugList_To_KeywordIssueDescription(List<Issue> issuelist)
        {
            List<StyleString> ret_style_string = new List<StyleString>();
            int processed_count = 0;

            foreach (Issue issue in issuelist)
            {
                List<StyleString> value_style_str = new List<StyleString>();
                String key = issue.Key;  // rd_comment_str = issue.comment;
                Boolean is_waived = false;

                processed_count++;
                if (String.IsNullOrWhiteSpace(key) != true)
                {
                    Color color_by_severity = Issue.KeywordIssue_report_FontColor;
                    if (issue.Status == Issue.STR_CLOSE)
                    {
                        color_by_severity = Issue.KeywordIssue_CLOSED_ISSUE_COLOR;
                    }
                    else if (issue.Status == Issue.STR_WAIVE)
                    {
                        color_by_severity = Issue.KeywordIssue_WAIVED_ISSUE_COLOR;
                        is_waived = true;
                    }
                    else // if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
                    {
                        switch (issue.Severity[0])
                        {
                            case 'A':
                                color_by_severity = Issue.KeywordIssue_A_ISSUE_COLOR;
                                break;
                            case 'B':
                                color_by_severity = Issue.KeywordIssue_B_ISSUE_COLOR;
                                break;
                            case 'C':
                                color_by_severity = Issue.KeywordIssue_C_ISSUE_COLOR;
                                break;
                            case 'D':
                                color_by_severity = Issue.KeywordIssue_D_ISSUE_COLOR;
                                break;
                            default:
                                // Use Default
                                break;
                        }

                    }

                    String str;
                    str = key + issue.Summary + "(" + issue.Severity + ")";
                    if (is_waived)
                    {
                        str += "(" + ReportGenerator.WAIVED_str + ")";
                    }
                    StyleString style_str = new StyleString(str, color_by_severity, Issue.KeywordIssue_report_Font,
                                                Issue.KeywordIssue_report_FontSize, Issue.KeywordIssue_report_FontStyle);
                    ret_style_string.Add(style_str);

                    // Add whole string into return_list
                    if (processed_count < issuelist.Count())
                    {
                        ret_style_string.Add(new StyleString("\n"));
                    }

                }
            }
            return ret_style_string;
        }

        //static public Dictionary<string, List<StyleString>> GenerateIssueDescription_Severity_by_Linked_Issue(List<Issue> issuelist)
        //{
        //    Dictionary<string, List<StyleString>> ret_list = new Dictionary<string, List<StyleString>>();

        //    foreach (Issue issue in issuelist)
        //    {
        //        List<StyleString> value_style_str = new List<StyleString>();
        //        String key = issue.Key;  // rd_comment_str = issue.comment;
        //        Boolean is_waived = false;

        //        if (key != "")
        //        {
        //            Color color_by_severity = Issue.ISSUE_DEFAULT_COLOR;
        //            if (issue.Status == Issue.STR_CLOSE)
        //            {
        //                color_by_severity = Issue.CLOSED_ISSUE_COLOR;
        //            }
        //            else if (issue.Status == Issue.STR_WAIVE)
        //            {
        //                color_by_severity = Issue.WAIVED_ISSUE_COLOR;
        //                is_waived = true;
        //            }
        //            else // if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
        //            {
        //                switch (issue.Severity[0])
        //                {
        //                    case 'A':
        //                        color_by_severity = ReportGenerator.LinkIssue_A_Issue_Color;
        //                        break;
        //                    case 'B':
        //                        color_by_severity = ReportGenerator.LinkIssue_B_Issue_Color;
        //                        break;
        //                    case 'C':
        //                        color_by_severity = ReportGenerator.LinkIssue_C_Issue_Color;
        //                        break;
        //                    case 'D':
        //                        color_by_severity = ReportGenerator.LinkIssue_D_Issue_Color;
        //                        break;
        //                    default:
        //                        // Use Default
        //                        break;
        //                }

        //            }

        //            String str;
        //            str = key + issue.Summary + "(" + issue.Severity + ")";
        //            if (is_waived)
        //            {
        //                str += "(" + KeywordReport.WAIVED_str + ")";
        //            }
        //            StyleString style_str = new StyleString(str, color_by_severity, ReportGenerator.LinkIssue_report_Font, 
        //                                        ReportGenerator.LinkIssue_report_FontSize, ReportGenerator.LinkIssue_report_FontStyle);
        //            value_style_str.Add(style_str);
        //            // Add whole string into return_list
        //            if (ret_list.ContainsKey(key))
        //            {
        //                continue; // shouldn't be here.
        //            }
        //            ret_list.Add(key, value_style_str);
        //        }
        //    }
        //    return ret_list;
        //}

        /*
        static public List<StyleString> BugList_To_LinkedIssueDescription(List<Issue> issuelist)
        {
            List<StyleString> ret_style_string = new List<StyleString>();
            int processed_count = 0;
            foreach (Issue issue in issuelist)
            {
                String key = issue.Key;  // rd_comment_str = issue.comment;
                processed_count++;

                //if (key != "")
                if (String.IsNullOrWhiteSpace(key) == false)
                {
                    List<StyleString> next_bug_description = Bug_To_LinkedIssueDescription(issue);
                    ret_style_string.AddRange(next_bug_description);
                    if (processed_count < issuelist.Count())
                    {
                        ret_style_string.Add(new StyleString("\n"));
                    }
                }
            }
            return ret_style_string;
        }
        */
        /*
        static public List<StyleString> Bug_To_LinkedIssueDescription(Issue issue)
        {
            List<StyleString> value_style_str = new List<StyleString>();
            String key = issue.Key;  // rd_comment_str = issue.comment;
            Boolean is_waived = false;

            if (String.IsNullOrWhiteSpace(key) == false)
            {
                Color color_by_severity = ReportGenerator.LinkIssue_report_FontColor;
                if (issue.Status == Issue.STR_CLOSE)
                {
                    color_by_severity = ReportGenerator.LinkIssue_CLOSED_ISSUE_COLOR;
                }
                else if (issue.Status == Issue.STR_WAIVE)
                {
                    color_by_severity = ReportGenerator.LinkIssue_WAIVED_ISSUE_COLOR;
                    is_waived = true;
                }
                else // if ((issue.Status != Issue.STR_CLOSE) && (issue.Status != Issue.STR_WAIVE))
                {
                    switch (issue.Severity[0])
                    {
                        case 'A':
                            color_by_severity = ReportGenerator.LinkIssue_A_Issue_Color;
                            break;
                        case 'B':
                            color_by_severity = ReportGenerator.LinkIssue_B_Issue_Color;
                            break;
                        case 'C':
                            color_by_severity = ReportGenerator.LinkIssue_C_Issue_Color;
                            break;
                        case 'D':
                            color_by_severity = ReportGenerator.LinkIssue_D_Issue_Color;
                            break;
                        default:
                            // Use Default
                            break;
                    }

                }

                String str;
                str = key + issue.Summary + "(" + issue.Severity + ")";
                if (is_waived)
                {
                    str += "(" + ReportGenerator.WAIVED_str + ")";
                }
                StyleString style_str = new StyleString(str, color_by_severity, ReportGenerator.LinkIssue_report_Font,
                                            ReportGenerator.LinkIssue_report_FontSize, ReportGenerator.LinkIssue_report_FontStyle);
                value_style_str.Add(style_str);
            }
            return value_style_str;
        }
        */
        /* 
        static public List<StyleString> TestCase_To_TestCaseSummary(TestCase testcasae)
        {
            List<StyleString> value_style_str = new List<StyleString>();
            String key = testcasae.Key;
            String summary = testcasae.Summary;

            if ((String.IsNullOrWhiteSpace(key) == false) || (String.IsNullOrWhiteSpace(summary) == false))
            {
                StyleString style_str = new StyleString(summary, default_color, default_font, default_size, default_fontstyle);
                value_style_str.Add(style_str);
            }

            return value_style_str;
        }
        */
        /*
        static public List<StyleString> TestCaseList_To_TestCaseSummary(List<TestCase> tc_list)
        {
            List<StyleString> ret_stylestring_list = new List<StyleString>();
            int processed_count = 0;
            foreach (TestCase tc in tc_list)
            {
                processed_count++;
                List<StyleString> next_testcase_summary = TestCase_To_TestCaseSummary(tc);
                ret_stylestring_list.AddRange(next_testcase_summary);
                if (processed_count < tc_list.Count())
                {
                    ret_stylestring_list.Add(new StyleString("\n"));
                }
            }
            return ret_stylestring_list;
        }
        */
    }
}
