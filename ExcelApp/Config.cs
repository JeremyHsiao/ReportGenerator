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
            @"workbook_BUG_Jira",              @".\SampleData\Jira 2022-09-03T10_48.xls", 
            @"workbook_TC_Jira",               @".\SampleData\TC_Jira 2022-09-03T11_07.xls",
            @"workbook_Summary",               @".\SampleData\Report_Template.xlsx",
            @"workbook_StandardTestReport",    @".\SampleData\TestFileFolder\0.0_DQA Test Report\BenQ_Medical_Standard Test Report.xlsx",
            @"Excel_Visible",                  @"true",
            @"Issue_Key_Prefix",               @"-",
            @"Issue_SheetName",                @"general_report",
            @"Issue_Row_NameDefine",           @"4",
            @"Issue_Row_DataBegin",            @"5",
            @"TC_Key_Prefix",                  @"TC",
            @"TC_SheetName",                   @"general_report",
            @"TC_Row_NameDefine",              @"4",
            @"TC_Row_DataBegin",               @"5",
            @"workbook_TC_Template",           @".\SampleData\TC_Jira_Template.xlsx",
        };

        public static String ReadAppSetting(string key)
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings; 
                String result  = appSettings[key];

                // if key not found --> return default value for those listed on the DefaultKeyValuePairList
                if (result == null)
                {
                    result = "";
                    for (int index = 0; index < DefaultKeyValuePairList.Count(); index += 2)
                    {
                        if (DefaultKeyValuePairList[index] == key)
                        {
                            result = DefaultKeyValuePairList[index + 1];
                            break;
                        }
                    }
                }
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                return "";
            }
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
