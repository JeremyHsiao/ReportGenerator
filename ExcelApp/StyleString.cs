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

        static public String StyleStringToString(StyleString stylestring)
        {
            return stylestring.text;
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
        static public List<StyleString> ExtendIssueDescription(string links_str, Dictionary<string, List<StyleString>> bug_list)
        {
            List<StyleString> extended_str = new List<StyleString>();

            // protection
            if ((links_str == null) || (bug_list == null)) return null;

            List<String> id_list = TestCase.Convert_LinksString_To_ListOfString(links_str);
            extended_str = ExtendIssueDescription(id_list, bug_list);
            return extended_str;
        }

        //
        // input: bug_id List
        // output: bug descriptions (one bug each line)
        //
        static public List<StyleString> ExtendIssueDescription(List<String> bug_id, Dictionary<string, List<StyleString>> bug_list)
        {
            List<StyleString> extended_str = new List<StyleString>();

            // protection
            if ((bug_id == null) || (bug_list == null)) return null;

            // replace each bug_id with full description seperated by newline and combine into one string
            StyleString new_line_str = new StyleString("\n");
            foreach (string key in bug_id)
            {
                string trimmed_key = key.Trim();
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

        static public void WriteStyleString(ref Range input_range, List<StyleString> sytle_string_list)
        {
            // Fill the text into excel cell with default font settings.
            string txt_str = "";
            foreach (StyleString style_str in sytle_string_list)
            {
                txt_str += style_str.Text;
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
                    // Font doesn't exist ==> no need to change font
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
            foreach (StyleString style_str in sytle_string_list)
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
                    if (fontTester.Name == StyleString.default_font)
                    {
                        // Font exists
                    }
                    else
                    {
                        // Font doesn't exist ==> no need to change font
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

        static public void WriteStyleString(Worksheet ws, int row, int col, List<StyleString> sytle_string_list)
        {
            Range input_range = ws.Cells[row,col];
            WriteStyleString(ref input_range, sytle_string_list);
        }

    }

}
