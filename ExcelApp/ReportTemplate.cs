﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace ExcelReportApplication
{
    class ReportTemplate
    {
    }

    public class ReportContentPair
    {
        public String Data;
        public int Row;
        public int Col;
        public String LabelData;
        public int LabelRow;
        public int LabelCol;
        public String VariableName;
        public Boolean On;
        public Boolean PreProcessing;
        public Boolean PostProcessing;

        public ReportContentPair(String label, String variableName)
        {
            InitialSetup(label, variableName);
        }
        public Boolean InitialSetup(String label, String variableName)
        {
            LabelData = label;
            VariableName = variableName;
            return !(String.IsNullOrWhiteSpace(LabelData) && String.IsNullOrWhiteSpace(VariableName));
        }

        public Boolean ToLabel(String label)
        {
            LabelData = label;
            return true;
        }

        public int Compare(String label)
        {
            return String.Compare(LabelData, label);
        }
    }

    public class ReportContent
    {
        public String Filename = "Report_FullFileName";
        public String TesterDataFilename = "TesterData_FullFileName";
        public String TemplateFilename = "Template_FullFileName";
        public String ExcelSheetName = "Report_Sheet_Name";
        public int TemplateEndRow = 22;
        public int TemplateEndCol = ExcelAction.ColumnNameToNumber("N");
        private List<ReportContentPair> content = new List<ReportContentPair>();
        public ReportContent(String filename, String testerDataFilename, String templateFilename, String excelSheetName, int templateEndRow, int templateEndCol)
        {
            Setup(filename, testerDataFilename, templateFilename, excelSheetName, templateEndRow, templateEndCol);
        }
        public void Setup(String filename, String testerDataFilename, String templateFilename, String excelSheetName, int templateEndRow, int templateEndCol)
        {
            Filename = filename;
            TesterDataFilename = testerDataFilename;
            TemplateFilename = templateFilename;
            ExcelSheetName = excelSheetName;
            TemplateEndRow = templateEndRow;
            TemplateEndCol = templateEndCol;
        }

        public Boolean ContainVariable(String variable)
        {
            foreach (ReportContentPair pair in content)
            {
                if (String.Compare(variable, pair.VariableName) == 0)
                {
                    return true;
                }
            }
            return false;
        }
    }

    static public class ReportManagement
    {
        static private int templateEndRow = 22;
        static private int templateEndCol = ExcelAction.ColumnNameToNumber("N");
        static private List<ReportContentPair> templateContentPair = new List<ReportContentPair>();
        static public String[] DefaultLabel = {
            "Model Name", "Panel Module" , "AD Board", "Smart BD / OS Version", "Speaker / AQ Version", "Test Stage", "Judgement", "Sample S/N", "Test by",
	        "Part No.", "T-Con Board", "Power Board", "Touch Sensor", "SW / PQ Version", "Test Period", "Approved by", 
            "Purpose:", "Condition:", "Equipment:", "Method:", "Criteria:", "Conclusion:", "Sample S/N:",
            // dummy label for label-less
            "TITLE" };
        static public String[] DefaultVariable = {
            "ModelName", "PanelModule" , "ADBoard", "SmartBD_OSVersion", "Speaker_AQVersion", "TestStage", "Judgement", "SampleSN", "Testby",
	        "PartNo", "TConBoard", "PowerBoard", "TouchSensor", "SW_PQVersion", "TestPeriod", "Approvedby", 
            "Purpose", "Condition", "Equipment", "Method", "Criteria", "Conclusion", "SampleSN",
            // Some variable are label-less
            "Title" };

        static public List<ReportContentPair> SetupHeaderTemplateContentPair(Worksheet worksheet, List<String> labelList, List<String> variableList)
        {
            List<String> variable_to_search = DefaultVariable.ToList();
            List<String> label_list = DefaultLabel.ToList();
            List<ReportContentPair> contentPairList = new List<ReportContentPair>();
            for (int row = 1; row <= templateEndRow; row++)
            {
                for (int col = 1; col <= templateEndCol; col++)
                {
                    String str = ExcelAction.GetCellTrimmedString(worksheet, row, col);
                    int strlen = str.Length;
                    if (strlen <= 3)
                        continue;

                    if ((str[0] != '$') || (str[strlen - 1] != '$'))
                        continue;

                    String variableName = (str.Substring(1, strlen - 2));
                    int index = variable_to_search.IndexOf(variableName);
                    if (index < 0)
                        continue;

                    variable_to_search.RemoveAt(index);
                    String label = label_list[index];
                    label_list.RemoveAt(index);

                    ReportContentPair cp = new ReportContentPair(label, variableName);
                    contentPairList.Add(cp);
                }
            }
            // search through excel until end of header section
            return contentPairList;
        }

    }

    /*
    public class ReportFormat
    {
        public String TemplateFilename = "Report_FullFileName";

        public List<int> RowHeight = new List<int>();
        public List<int> ColumnWidth = new List<int>();
    }

    public class ReportData
    {
        public String Filename = "Report_FullFileName";
        public String TesterDataFilename = "TesterData_FullFileName";
        public String ExcelSheetName = "Report_Sheet_Name";
        public String Title = "Report_Name";
        public String ModelName = "Model Name";
        public String PartNo = "Part_No";
        public String PanelModule = "Panel_Module";
        public String TCONBoard = "T-Con_Board";
        public String ADBoard = "AD_Board";
        public String PowerBoard = "Power_Board";
        public String SmartBD_OSVersion = "Smart_BD_OS_Version";
        public String TouchSensor = "Touch_Sensor";
        public String Speaker_AQVersion = "Speaker_AQ_Version";
        public String SW_PQVersion = "SW_PQ_Version";
        public String TestStage = " ";
        public String TestQTY_SN = " ";
        public String SampleSN = " ";
        public String TestPeriodBegin = "2023/07/10";
        public String TestPeriodEnd = "2023/07/10";
        public String TestPeriodFormat = "yyyy/MM/dd ~ yyyy/MM/dd";
        public String Judgement = " ";
        public String Tested_by = " ";
        public String Approved_by = "Jeremy Hsiao";
        public String Purpose = "Test Purpose";
        public String Condition = "Test Condition";
        public String Equipment = "Test Equipment";
        public String Method = "Test Method";
        public String Criteria = "Test Criteria";
        public String Conclusion = " ";

        public int Title_row = 1;
        public int ModelName_row = 3;
        public int PartNo_row = 3;
        public int PanelModule_row = 4;
        public int TCONBoard_row = 4;
        public int ADBoard_row = 5;
        public int PowerBoard_row = 5;
        public int SmartBD_OSVersion_row = 6;
        public int TouchSensor_row = 6;
        public int Speaker_AQVersion_row = 7;
        public int SW_PQVersion_row = 7;
        public int TestStage_row = 8;
        public int TestQTY_SN_row = 8;
        public int SampleSN_row = 8;
        public int TestPeriodBegin_row = 8;
        public int TestPeriodEnd_row = 8;
        public int Judgement_row = 9;
        public int Tested_by_row = 9;
        public int Approved_by_row = 9;
        public int Purpose_row = 12;
        public int Condition_row = 14;
        public int Equipment_row = 16;
        public int Method_row = 18;
        public int Criteria_row = 20;
        public int Conclusion_row = 22;

        public int Title_col = ExcelAction.ColumnNameToNumber("A");
        public int ModelName_col = ExcelAction.ColumnNameToNumber("D");
        public int PartNo_col = ExcelAction.ColumnNameToNumber("J");
        public int PanelModule_col = ExcelAction.ColumnNameToNumber("D");
        public int TCONBoard_col = ExcelAction.ColumnNameToNumber("J");
        public int ADBoard_col = ExcelAction.ColumnNameToNumber("D");
        public int PowerBoard_col = ExcelAction.ColumnNameToNumber("J");
        public int SmartBD_OSVersion_col = ExcelAction.ColumnNameToNumber("D");
        public int TouchSensor_col = ExcelAction.ColumnNameToNumber("J");
        public int Speaker_AQVersion_col = ExcelAction.ColumnNameToNumber("D");
        public int SW_PQVersion_col = ExcelAction.ColumnNameToNumber("J");
        public int TestStage_col = ExcelAction.ColumnNameToNumber("D");
        public int TestQTY_SN_col = ExcelAction.ColumnNameToNumber("H");
        public int SampleSN_col = ExcelAction.ColumnNameToNumber("H");
        public int TestPeriodBegin_col = ExcelAction.ColumnNameToNumber("L");
        public int TestPeriodEnd_col = ExcelAction.ColumnNameToNumber("L");
        public int Judgement_col = ExcelAction.ColumnNameToNumber("D");
        public int Tested_by_col = ExcelAction.ColumnNameToNumber("H");
        public int Approved_by_col = ExcelAction.ColumnNameToNumber("L");
        public int Purpose_col = ExcelAction.ColumnNameToNumber("C");
        public int Condition_col = ExcelAction.ColumnNameToNumber("C");
        public int Equipment_col = ExcelAction.ColumnNameToNumber("C");
        public int Method_col = ExcelAction.ColumnNameToNumber("C");
        public int Criteria_col = ExcelAction.ColumnNameToNumber("C");
        public int Conclusion_col = ExcelAction.ColumnNameToNumber("C");
    }
    */

}