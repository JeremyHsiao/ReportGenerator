using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Text.RegularExpressions;

namespace ExcelReportApplication
{
    static public class HeaderTemplate
    {
        static public int StartRow = 1;
        static public int EndRow = 9;
        static public int StartCol = 1;
        static public int EndCol = ExcelAction.ColumnNameToNumber('N');

        static public String Variable_ReportFileName = "$FileName$";
        static public String Variable_ReportSheetName = "$SheetName$";
        static public String Variable_Assignee = "$Assignee$";
        static public String Variable_Today = "$Today$";
        static public String Variable_TC_LinkedIssue = "$LinkedIssue$";
        static public String Variable_KEEP = "$KEEP$";

        static private String ReportFileName = "$FileName$";
        static private String ReportSheetName = "$SheetName$";
        static private String Assignee = "$Assignee$";
        static private String Today = "$Today$";
        static private List<StyleString> TC_LinkedIssue = StyleString.StringToListOfStyleString(Variable_TC_LinkedIssue);

        static public void ResetVariables()
        {
            ReportFileName = Variable_ReportFileName;
            ReportSheetName = Variable_ReportSheetName;
            Assignee = Variable_Assignee;
            Today = Variable_Today;
            TC_LinkedIssue = StyleString.StringToListOfStyleString(Variable_TC_LinkedIssue);
        }

        static public void UpdateVariables_FilenameSheetname(String filename, String sheetname)
        {
            ReportFileName = filename;
            ReportSheetName = sheetname;
        }

        static public void UpdateVariables_TodayAssigneeLinkedIssue(String today, String assignee, List<StyleString> LinkedIssue)
        {
            Assignee = assignee;
            Today = today;
            TC_LinkedIssue = LinkedIssue;
        }

        //static public void UpdateVariables(String filename = "", String sheetname = "", String assignee = "", String today = "", List<StyleString> LinkedIssue = null)
        //{
        //    if (String.IsNullOrWhiteSpace(filename) == false)
        //    {
        //        ReportFileName = filename;
        //    }
        //    if (String.IsNullOrWhiteSpace(sheetname) == false)
        //    {
        //        ReportSheetName = sheetname;
        //    }
        //    if (String.IsNullOrWhiteSpace(assignee) == false)
        //    {
        //        Assignee = assignee;
        //    }
        //    if (String.IsNullOrWhiteSpace(today) == false)
        //    {
        //        Today = today;
        //    }
        //    if (LinkedIssue != null)
        //    {
        //        TC_LinkedIssue = LinkedIssue;
        //    }
        //}

        static private Boolean CheckAndReplace(Worksheet ws, int row, int col, String from, String to)
        {
            Boolean b_ret = false;
            if (ExcelAction.GetCellValue(ws, row, col) != null)
            {
                String to_check = ExcelAction.GetCellValue(ws, row, col).ToString();
                if (to_check.Contains(from))
                {
                    ExcelAction.ReplaceText(ws, row, col, from, to);
                    b_ret = true;
                }
            }
            else
            {
                b_ret = true;
            }
            return b_ret;
        }

        static private Boolean CheckAndReplaceConclusion(Worksheet ws, int row, int col, String from, List<StyleString> to)
        {
            Boolean b_ret = false;
            if (ExcelAction.GetCellValue(ws, row, col) != null)
            {
                String to_check = ExcelAction.GetCellValue(ws, row, col).ToString();
                if (to_check.Contains(Variable_TC_LinkedIssue))
                {
                    StyleString.WriteStyleString(ws, row, col, TC_LinkedIssue);
                }
                b_ret = true;
            }
            else
            {
                b_ret = true;
            }
            return b_ret;
        }

        static public Boolean ReplaceHeaderVariableWithValue(Worksheet report_worksheet)
        {
            Boolean b_ret = false;

            for (int row_index = StartRow; row_index <= EndRow; row_index++)
            {
                for (int col_index = StartCol; col_index <= EndCol; col_index++)
                {
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_ReportFileName, ReportFileName);
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_ReportSheetName, ReportSheetName);
                    Assignee = Regex.Replace(Assignee, "[\u4E00-\u9FFF]", ""); // 移除中文
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_Assignee, Assignee);
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_Today, Today);
                    CheckAndReplaceConclusion(report_worksheet, row_index, col_index, Variable_TC_LinkedIssue, TC_LinkedIssue);
                }
            }
            b_ret = true;
            return b_ret;
        }

        static public Boolean CopyAndUpdateHeader(Worksheet template_worksheet, Worksheet report_worksheet)
        {
            Boolean b_ret = false;

            //ExcelAction.CopyRowHeight(template_worksheet, report_worksheet, StartRow, EndRow);
            //ExcelAction.CopyColumnWidth(template_worksheet, report_worksheet, StartCol, EndCol);
            ExcelAction.CopyPasteRows(template_worksheet, report_worksheet, StartRow, EndRow);

            b_ret = ReplaceHeaderVariableWithValue(report_worksheet);
            return b_ret;
        }

        static private List<int> KEEP_ROW = new List<int>(), KEEP_COL = new List<int>();
        static private List<Object> KEEP_CELL = new List<Object>();

        static public Boolean CopyKEEPCell(Worksheet template_worksheet, Worksheet report_workshee)
        {
            Boolean b_ret = false;
            KEEP_ROW.Clear();
            KEEP_COL.Clear();
            KEEP_CELL.Clear();
            for (int row_index = StartRow; row_index <= EndRow; row_index++)
            {
                for (int col_index = StartCol; col_index <= EndCol; col_index++)
                {
                    Object obj = ExcelAction.GetCellValue(template_worksheet, row_index, col_index);
                    if (obj != null)
                    {
                        String to_check = obj.ToString();
                        if (to_check.Contains(Variable_KEEP))
                        {
                            KEEP_ROW.Add(row_index);
                            KEEP_COL.Add(col_index);
                            obj = ExcelAction.GetCellValue(report_workshee, row_index, col_index);
                            KEEP_CELL.Add(obj);
                        }
                    }
                }
            }
            return b_ret;
        }

        static public Boolean CopyAndUpdateHeader_with_KEEP(Worksheet template_worksheet, Worksheet report_worksheet)
        {
            Boolean b_ret = false;

            CopyKEEPCell(template_worksheet, report_worksheet);
            //ExcelAction.CopyRowHeight(template_worksheet, report_worksheet, StartRow, EndRow);
            //ExcelAction.CopyColumnWidth(template_worksheet, report_worksheet, StartCol, EndCol);
            ExcelAction.CopyPasteRows(template_worksheet, report_worksheet, StartRow, EndRow);

            b_ret = ReplaceHeaderVariableWithValue(report_worksheet);

            // PasteKEEPCell
            int count = KEEP_ROW.Count();
            while (count-- > 0)
            {
                int row = KEEP_ROW[count], col = KEEP_COL[count];
                Object obj = KEEP_CELL[count];
                ExcelAction.SetCellValue(report_worksheet, row, col, obj);
            }

            return b_ret;
        }

    }

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
            "ModelName", "PanelModule" , "ADBoard", "SmartBD_OSVersion", "Speaker_AQVersion", "TestStage", "Judgement", "SampleSN", "Assignee",
	        "PartNo", "TConBoard", "PowerBoard", "TouchSensor", "SW_PQVersion", "TestPeriod", "Approvedby", 
            "Purpose", "Condition", "Equipment", "Method", "Criteria", "Conclusion", "SampleSN",
            // Some variable are label-less
            "Filename" };

        static public List<ReportContentPair> SetupHeaderTemplateContentPair(Worksheet worksheet, List<String> labelList, List<String> variableList)
        {
            List<String> variable_to_search = new List<String>();
            List<String> label_list_to_search = new List<String>(); 
            List<ReportContentPair> contentPairList = new List<ReportContentPair>();
            variable_to_search.AddRange(variableList);
            label_list_to_search.AddRange(labelList);
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
                    String label = label_list_to_search[index];
                    label_list_to_search.RemoveAt(index);

                    ReportContentPair cp = new ReportContentPair(label, variableName);
                    contentPairList.Add(cp);
                }
            }
            // search through excel until end of header section
            return contentPairList;
        }

        static public Boolean ReadHeaderVariablesAccordingToTemplate(String input_excel_file)
        {
            // Open Header Template Excel workbook
            Workbook wb_input_excel = ExcelAction.OpenExcelWorkbook(filename: input_excel_file, ReadOnly: true);
            if (wb_input_excel == null)
            {
                LogMessage.WriteLine("ERR: Open workbook failed in ReadHeaderVariablesAccordingToTemplate(): " + input_excel_file);
                return false;
            }

            Worksheet ws_input_excel;
            String sheet_to_select = InputExcel.SheetName_HeaderTemplate_Source;
            if (ExcelAction.WorksheetExist(wb_input_excel, sheet_to_select) == false)
            {
                LogMessage.WriteLine("ERR: template worksheet doesn't exist on excel: " + input_excel_file);
                return false;
            }
            ws_input_excel = ExcelAction.Find_Worksheet(wb_input_excel, sheet_to_select);

            List<ReportContentPair> rcp = SetupHeaderTemplateContentPair(ws_input_excel, DefaultLabel.ToList(), DefaultVariable.ToList());

            ExcelAction.CloseExcelWorkbook(wb_input_excel);

            // Template variable has been read, start to read all reports

            return true;
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
