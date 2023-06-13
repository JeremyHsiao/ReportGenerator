using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Collections.Specialized;

namespace ExcelReportApplication
{

    class XMLConfig
    {
        private static String[] DefaultKeyValuePairList =
        {
            @"workbook_BUG_Jira",                       @".\SampleData\Jira 2022-09-03T10_48.xls", 
            @"workbook_TC_Jira",                        @".\SampleData\TC_Jira 2022-09-03T11_07.xls",
            @"workbook_Summary",                        @".\SampleData\Report_Template.xlsx",
            @"workbook_StandardTestReport",             @".\SampleData\TestFileFolder\0.0_DQA Test Report\BenQ_Medical_Standard Test Report.xlsx",
            @"Keyword_default_report_dir",              @".\SampleData\EVT_Winnie_Keyword2.5_keyword\All function",
            @"Excel_Visible",                           @"true",
            @"Issue_Key_Prefix",                        @"-",
            @"Issue_SheetName",                         @"general_report",
            @"Issue_Row_NameDefine",                    @"4",
            @"Issue_Row_DataBegin",                     @"5",
            @"TC_Key_Prefix",                           @"TC",
            @"TC_SheetName",                            @"general_report",
            @"TC_Row_NameDefine",                       @"4",
            @"TC_Row_DataBegin",                        @"5",
            @"workbook_TC_Template",                    @".\SampleData\TC_Jira_Template.xlsx",
            @"LinkIssueFilterStatusString",             @"Close (0), Waive (0.1)",
            @"TestReport_Row_UserStart",                @"27",
            @"TestReport_Column_Keyword_Indentifier",   @"2",
            @"TestReport_Regex_Keyword_Indentifier",    @"(?i)Item",
            @"TestReport_Column_Keyword_Location",      @"3",
            @"TestReport_Result_Titlle_Offset_Row",     @"1",
            @"TestReport_Result_Titlle_Offset_Col",     @"1",
            @"TestReport_BugStatus_Titlle_Offset_Row",  @"1",
            @"TestReport_BugStatus_Titlle_Offset_Col",  @"3",
            @"TestReport_BugList_Titlle_Offset_Row",    @"2",
            @"TestReport_BugList_Titlle_Offset_Col",    @"1",
            @"TestReport_Default_Output_Dir",           @"",
            @"default_report_Font",                     @"Calibri",
            @"default_report_FontSize",                 @"10",
            @"default_report_FontColor",                @"Black",
            @"default_report_FontStyle",                @"Regular",
            //@"buglist_sorting",                       @"Severity,Key",
            //@"buglist_Severity_A",                    @"Red",
            //@"buglist_Severity_B",                    @"DarkOrange",
            //@"buglist_Severity_C",                    @"Black",
            //@"buglist_Severity_D",                    @"Black",
            //@"buglist_Status_Waive",                  @"Black",
            //@"buglist_Status_Close",                  @"Black",
            //@"buglist_Status_Other",                  @"Black",
            @"Keyword_report_A_Issue_Color",            @"Red",
            @"Keyword_report_B_Issue_Color",            @"DarkOrange",
            @"Keyword_report_C_Issue_Color",            @"Black",
            @"Keyword_Auto_Correct_Worksheet",          @"false",

        };

        // buglist_Severity support: Color.xxxxx
        // buglist_Status_Other support: 

        public static String GetDefaultValue(String key)
        {
            String result = "";

            for (int index = 0; index < DefaultKeyValuePairList.Count(); index += 2)
            {
                if (DefaultKeyValuePairList[index] == key)
                {
                    result = DefaultKeyValuePairList[index + 1];
                    return result;
                }
            }
            return result;
        }

        // make is private so that all ReadAppSettings are type-checked.
        private static String ReadAppSetting(String key)
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings; 
                String result  = appSettings[key];

                // if key not found --> return default value for those listed on the DefaultKeyValuePairList
                if (result == null) return GetDefaultValue(key);
                else return result;
            }
            catch (ConfigurationErrorsException)
            {
                return GetDefaultValue(key);
            }
        }

        public static String ReadAppSetting_String(String key)
        {
            return ReadAppSetting(key);
        }

        public static Boolean ReadAppSetting_Boolean(String key)
        {
            Boolean ret_value;
            if (!Boolean.TryParse(ReadAppSetting(key), out ret_value))
            {
                // TryParse failed, use default value
                if (!Boolean.TryParse(GetDefaultValue(key), out ret_value))
                {
                    // If still failed, it indicated that default value in code is not correct!
                    // highlight error
                    Console.WriteLine("Default value for " + key + " is not a Boolean. Please check");
                }
            }
            return ret_value;
        }

        public static int ReadAppSetting_int(String key)
        {
            int ret_value;
            if (!int.TryParse(ReadAppSetting(key), out ret_value))
            {
                // TryParse failed, use default value
                if (!int.TryParse(GetDefaultValue(key), out ret_value))
                {
                    // If still failed, it indicated that default value in code is not correct!
                    // highlight error
                    Console.WriteLine("Default value for " + key + " is not an int. Please check");
                }
            }
            return ret_value;
        }

        public static System.Drawing.Color ReadAppSetting_Color(String key)
        {
            System.Drawing.Color ret_value;
            int int_value;
            string input_str = ReadAppSetting(key);
            // Because color could be a string (ex:Black) or a ARGB value #00000000
            if (int.TryParse(input_str, out int_value))
            {
                ret_value = System.Drawing.Color.FromArgb(int_value);
            }
            else
            {
                //  Treat string as color name.
                //  If the name parameter is not the valid name of a predefined color, 
                //  the FromName method creates a Color structure that has an ARGB value of 0 (that is, all ARGB components are 0).
                ret_value = System.Drawing.Color.FromName(input_str);
            }
            return ret_value;
        }

        public static System.Drawing.FontStyle ReadAppSetting_FontStyle(String key)
        {
            System.Drawing.FontStyle ret_value;
            int int_value;
            string input_str = ReadAppSetting(key);
            // Because FontStyle could be a string or a value
            if (int.TryParse(input_str, out int_value))
            {
                ret_value = (System.Drawing.FontStyle)int_value;
            }
            else
            {
                //  Treat string as FontStyle name.
                ret_value = (System.Drawing.FontStyle)Enum.Parse((typeof(System.Drawing.FontStyle)), input_str);
            }
            return ret_value;
        }

        public static void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }  

    }
}
