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
        // Updated to format 002
        private static String[] DefaultKeyValuePairList =
        {
            @"CONFIG_FORMAT_VERSION",                   @"000", // DO NOT CHANGE THIS DEFAULT "000" -- all previous format are treated as 000
            @"CONFIG_DEFAULT_VALUE_VERSION",            @"000", // DO NOT CHANGE THIS DEFAULT "000" -- all previous default value are treated as 000
            @"workbook_BUG_Jira",                       @".\SampleData\Jira 2022-09-03T10_48.xls", 
            @"workbook_TC_Jira",                        @".\SampleData\TC_Jira 2022-09-03T11_07.xls",
            @"Keyword_default_report_dir",              @".\SampleData\EVT_Winnie_Keyword2.5_keyword\All function",
            @"workbook_TC_Template",                    @".\SampleData\TC_Jira_Template.xlsx",
            @"Report_E_Remove_AUO_Sheet_Default_Excel", @".\SampleData\EVT_Winnie_Keyword2.5_keyword\Copy_Report_Excel_List.xlsx",
            @"Report_D_Copy_Only_Default_Excel",        @".\SampleData\EVT_Winnie_Keyword2.5_keyword\Copy_Report_Excel_List.xlsx",
            @"Report_C_Default_Excel",                  @".\SampleData\EVT_Winnie_Keyword2.5_keyword\Copy_Report_Excel_List.xlsx",
            @"Report_A_Default_Excel",                  @".\SampleData\EVT_Winnie_Keyword2.5_keyword\Copy_Report_Excel_List.xlsx",
            @"InputExcel_Sheetname_ReportList",         @"ReportList",
            @"InputExcel_Sheetname_HeaderTemplate",     @"BeforeLine21",
            @"InputExcel_Sheetname_Bug",                @"Bug",
            @"InputExcel_Sheetname_TestCase",           @"TestCase",
            @"Report_C_CopyFileOnly",                   @"false",
            @"Report_C_Remove_AUO_Internal",            @"false",
            @"Report_C_Update_Report_Sheetname",        @"true",
            @"Report_C_Clear_Keyword_Result",           @"false",
            @"Report_C_Hide_Keyword_Result_Bug_Row",    @"false",
            @"Report_C_Replace_Conclusion",             @"false",
            @"Report_C_Update_Header_by_Template",      @"false",
            @"Report_C_Update_Conclusion",              @"false",
            @"Report_C_Update_Judgement",               @"false",
            @"Report_C_Update_Sample_SN",               @"false", 
            @"Report_C_SampleSN_String",                @"Refer to DUT_Allocation_Matrix table",
            // lagacy options - BEGIN
            @"Report_C_Update_Full_Header",             @"false",
            @"Update_Report_Title_by_Sheetname",        @"false",
            @"Update_Model_Name",                       @"false",
            @"Update_Part_No",                          @"false",
            @"Update_Panel_Module",                     @"false",
            @"Update_TCON_Board",                       @"false",
            @"Update_AD_Board",                         @"false",
            @"Update_Power_Board",                      @"false",
            @"Update_Smart_BD_OS_Version",              @"false",
            @"Update_Touch_Sensor",                     @"false",
            @"Update_Speaker_AQ_Version",               @"false",
            @"Update_SW_PQ_Version",                    @"false",
            @"Update_Test_Stage",                       @"false",
            @"Update_Test_QTY_SN",                      @"false",
            @"Update_Test_Period_Begin",                @"false",
            @"Update_Test_Period_End",                  @"false",
            @"Update_Judgement",                        @"false",
            @"Update_Tested_by",                        @"false",
            @"Update_Approved_by",                      @"false",
            @"Default_Model_Name",                      @"Model Name",
            @"Default_Part_No",                         @"Part_No",
            @"Default_Panel_Module",                    @"Panel_Module",
            @"Default_TCON_Board",                      @"T-Con_Board",
            @"Default_AD_Board",                        @"AD_Board",
            @"Default_Power_Board",                     @"Power_Board",
            @"Default_Smart_BD_OS_Version",             @"Smart_BD_OS_Version",
            @"Default_Touch_Sensor",                    @"Touch_Sensor",
            @"Default_Speaker_AQ_Version",              @"Speaker_AQ_Version",
            @"Default_SW_PQ_Version",                   @"Default_SW_PQ_Version",
            @"Default_Test_Stage",                      @"DVT",
            @"Default_Test_QTY_SN",                     @"QTY_or_SN",
            @"Default_Test_Period_Begin",               @"2023/07/10",
            @"Default_Test_Period_End",                 @"2023/07/19",
            @"Default_Judgement",                       @" ",
            @"Default_Tested_by",                       @" ", 
            @"Default_Approved_by",                     @"Jeremy Hsiao",
            @"Report_C_ReadHeaderItem",                 @"false",
            // lagacy options - END
            @"Excel_Visible",                           @"true",
            @"Issue_Key_Prefix",                        @"-",
            @"BugList_ExportedSheetName",               @"general_report",
            @"Issue_Row_NameDefine",                    @"4",
            @"Issue_Row_DataBegin",                     @"5",
            @"TC_Key_Prefix",                           @"TC",
            @"TCList_ExportedSheetName",                @"general_report",
            @"TC_SheetName",                            @"general_report",
            @"TC_Row_NameDefine",                           @"4",
            @"TC_Row_DataBegin",                            @"5",
            @"LinkIssueFilterStatusString",                 @"Close (0)",      // @"Close (0), Waive (0.1)",
            @"KeywordIssue_Row_UserStart",                  @"27",
            @"KeywordIssue_Column_Keyword_Indentifier",     @"2",
            @"KeywordIssue_Regex_Keyword_Indentifier",      @"(?i)Item",
            @"KeywordIssue_Column_Keyword_Location",        @"3",
            @"KeywordIssue_Result_Title_Offset_Row",       @"1",
            @"KeywordIssue_Result_Title_Offset_Col",       @"1",
            @"KeywordIssue_BugStatus_Title_Offset_Row",    @"1",
            @"KeywordIssue_BugStatus_Title_Offset_Col",    @"3",
            @"KeywordIssue_BugList_Title_Offset_Row",      @"2",
            @"KeywordIssue_BugList_Title_Offset_Col",      @"1",
            @"TestReport_Default_Output_Dir",               @"",
            @"TestReport_Default_Source_Path",              @"",
            @"KeywordIssueFilterStatusString",              @"Close (0)",
            @"TestReport_Default_Judgement",                @"N/A",
            @"TestReport_Default_Conclusion",               @" ",
            @"default_report_Font",                         @"Calibri",
            @"default_report_FontSize",                     @"10",
            @"default_report_FontColor",                    @"Black",
            @"default_report_FontStyle",                @"Regular",
            //@"buglist_sorting",                       @"Severity,Key",
            //@"buglist_Severity_A",                    @"Red",
            //@"buglist_Severity_B",                    @"DarkOrange",
            //@"buglist_Severity_C",                    @"Black",
            //@"buglist_Severity_D",                    @"Black",
            //@"buglist_Status_Waive",                  @"Black",
            //@"buglist_Status_Close",                  @"Black",
            //@"buglist_Status_Other",                  @"Black",
            @"LinkIssue_report_Font",                   @"Calibri",
            @"LinkIssue_report_FontSize",               @"10",
            @"LinkIssue_report_FontColor",              @"Black",
            @"LinkIssue_report_FontStyle",              @"Regular",
            @"LinkIssue_A_Issue_Color",                 @"Red",
            @"LinkIssue_B_Issue_Color",                 @"Black",
            @"LinkIssue_C_Issue_Color",                 @"Black",
            @"LinkIssue_D_Issue_Color",                 @"Black",

            @"KeywordIssue_report_Font",                @"Calibri",
            @"KeywordIssue_report_FontSize",            @"10",
            @"KeywordIssue_report_FontColor",           @"Black",
            @"KeywordIssue_report_FontStyle",           @"Regular",
            @"KeywordIssue_A_Issue_Color",              @"Red",
            @"KeywordIssue_B_Issue_Color",              @"Black",
            @"KeywordIssue_C_Issue_Color",              @"Black",
            @"KeywordIssue_D_Issue_Color",              @"Black",
            @"KeywordIssue_report_replace_conclusion",  @"false",
            @"KeywordIssue_report_Hide_Result_Bug",     @"false",  
            @"KeywordIssue_report_Correct_Worksheet",   @"false",
            //
            // report 2
            @"workbook_Summary",                        @".\SampleData\Report_Template.xlsx",
            // report 3
            @"workbook_StandardTestReport",             @".\SampleData\TestFileFolder\0.0_DQA Test Report\BenQ_Medical_Standard Test Report.xlsx",
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
                String result = appSettings[key];

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
                    LogMessage.WriteLine("Default value for " + key + " is not a Boolean. Please check");
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
                    LogMessage.WriteLine("Default value for " + key + " is not an int. Please check");
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
                LogMessage.WriteLine("Error writing app settings");
            }
        }

    }
}
