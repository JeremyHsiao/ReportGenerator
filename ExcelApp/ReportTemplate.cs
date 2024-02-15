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

        static public void UpdateVariables_TodayAssignee(String today, String assignee)
        {
            Assignee = assignee;
            Today = today;
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

        static private List<String> VariableName = new List<String>();
        static private List<int> VariableRow = new List<int>();
        static private List<int> VariableCol = new List<int>();

        static public Boolean ReplaceHeaderVariableWithValue(Worksheet report_worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            Boolean b_ret = false;

            for (int row_index = startRow; row_index <= endRow; row_index++)
            {
                for (int col_index = startCol; col_index <= endCol; col_index++)
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

        static public Boolean ReplaceHeaderVariableWithValue(Worksheet report_worksheet)
        {
            return ReplaceHeaderVariableWithValue(report_worksheet, StartRow, StartCol, EndRow, EndCol);
        }

        static public Boolean FindHeaderVariableLocation(Worksheet template_worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            Boolean b_ret = false;
            VariableName.Clear();
            VariableRow.Clear();
            VariableCol.Clear();
            for (int row_index = startRow; row_index <= endRow; row_index++)
            {
                for (int col_index = startCol; col_index <= endCol; col_index++)
                {
                    Object obj = ExcelAction.GetCellValue(template_worksheet, row_index, col_index);
                    if (obj != null)
                    {
                        String to_check = obj.ToString();
                        if (to_check.Contains(Variable_ReportFileName))
                        {
                            VariableName.Add(Variable_ReportFileName);
                            VariableRow.Add(row_index);
                            VariableCol.Add(col_index);
                        }
                        if (to_check.Contains(Variable_ReportSheetName))
                        {
                            VariableName.Add(Variable_ReportSheetName);
                            VariableRow.Add(row_index);
                            VariableCol.Add(col_index);
                        }
                        if (to_check.Contains(Variable_Assignee))
                        {
                            VariableName.Add(Variable_Assignee);
                            VariableRow.Add(row_index);
                            VariableCol.Add(col_index);
                        }
                        if (to_check.Contains(Variable_Today))
                        {
                            VariableName.Add(Variable_Today);
                            VariableRow.Add(row_index);
                            VariableCol.Add(col_index);
                        }
                        if (to_check.Contains(Variable_TC_LinkedIssue))
                        {
                            VariableName.Add(Variable_TC_LinkedIssue);
                            VariableRow.Add(row_index);
                            VariableCol.Add(col_index);
                        }
                    }
                }
            }
            b_ret = true;
            return b_ret;
        }

        static public Boolean PasteHeaderVariableContent(Worksheet report_worksheet)
        {
            Boolean b_ret = false;
            // PasteKEEPCell
            int count = VariableName.Count();
            while (count-- > 0)
            {
                int row_index = VariableRow[count], col_index = VariableCol[count];
                String variableName = VariableName[count];
                if (variableName == Variable_ReportFileName)
                {
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_ReportFileName, ReportFileName);
                }
                if (variableName == Variable_ReportSheetName)
                {
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_ReportSheetName, ReportSheetName);
                }
                if (variableName == Variable_Assignee)
                {
                    Assignee = Regex.Replace(Assignee, "[\u4E00-\u9FFF]", ""); // 移除中文
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_Assignee, Assignee);
                }
                if (variableName == Variable_Today)
                {
                    CheckAndReplace(report_worksheet, row_index, col_index, Variable_Today, Today);
                }
                if (variableName == Variable_TC_LinkedIssue)
                {
                    CheckAndReplaceConclusion(report_worksheet, row_index, col_index, Variable_TC_LinkedIssue, TC_LinkedIssue);
                }
            }
            b_ret = true;
            return b_ret;
        }

        
        /*
        static public Boolean CopyAndUpdateHeader(Worksheet template_worksheet, Worksheet report_worksheet)
        {
            Boolean b_ret = false;

            //ExcelAction.CopyRowHeight(template_worksheet, report_worksheet, StartRow, EndRow);
            //ExcelAction.CopyColumnWidth(template_worksheet, report_worksheet, StartCol, EndCol);
            ExcelAction.CopyPasteRows(template_worksheet, report_worksheet, StartRow, EndRow);

            b_ret = ReplaceHeaderVariableWithValue(report_worksheet);
            return b_ret;
        }
*/
        static private List<int> KEEP_ROW = new List<int>(), KEEP_COL = new List<int>();
        static private List<Object> KEEP_CELL = new List<Object>();

        static public Boolean CopyKEEPCell(Worksheet template_worksheet, Worksheet report_worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            Boolean b_ret = false;
            KEEP_ROW.Clear();
            KEEP_COL.Clear();
            KEEP_CELL.Clear();
            for (int row_index = startRow; row_index <= endRow; row_index++)
            {
                for (int col_index = startCol; col_index <= endCol; col_index++)
                {
                    Object obj = ExcelAction.GetCellValue(template_worksheet, row_index, col_index);
                    if (obj != null)
                    {
                        String to_check = obj.ToString();
                        if (to_check.Contains(Variable_KEEP))
                        {
                            KEEP_ROW.Add(row_index);
                            KEEP_COL.Add(col_index);
                            obj = ExcelAction.GetCellValue(report_worksheet, row_index, col_index);
                            KEEP_CELL.Add(obj);
                        }
                    }
                }
            }
            b_ret = true;
            return b_ret;
        }

        static public Boolean CopyKEEPCell(Worksheet template_worksheet, Worksheet report_worksheet)
        {
            return CopyKEEPCell(template_worksheet, report_worksheet, StartRow, StartCol, EndRow, EndCol);
        }

        static public Boolean PasteKEEPCell(Worksheet worksheet)
        {
            Boolean b_ret = false;
            // PasteKEEPCell
            int count = KEEP_ROW.Count();
            while (count-- > 0)
            {
                int row = KEEP_ROW[count], col = KEEP_COL[count];
                Object obj = KEEP_CELL[count];
                ExcelAction.SetCellValue(worksheet, row, col, obj);
            }
            b_ret = true;
            return b_ret;
        }

        static public Boolean CopyAndUpdateHeader_with_KEEP(Worksheet template_worksheet, Worksheet report_worksheet, int startRow, int endRow)
        {
            Boolean b_ret = false;

            CopyKEEPCell(template_worksheet, report_worksheet);
            //ExcelAction.CopyRowHeight(template_worksheet, report_worksheet, StartRow, EndRow);
            //ExcelAction.CopyColumnWidth(template_worksheet, report_worksheet, StartCol, EndCol);
            ExcelAction.CopyPasteRows(template_worksheet, report_worksheet, startRow, endRow);

            b_ret = ReplaceHeaderVariableWithValue(report_worksheet);

            PasteKEEPCell(report_worksheet);

            return b_ret;
        }

        static public Boolean CopyAndUpdateHeader_with_KEEP(Worksheet template_worksheet, Worksheet report_worksheet)
        {
            return CopyAndUpdateHeader_with_KEEP(template_worksheet, report_worksheet, StartRow, EndRow);
        }

        static public Boolean FindKEEPCellLocation(Worksheet template_worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            Boolean b_ret = false;
            KEEP_ROW.Clear();
            KEEP_COL.Clear();
            KEEP_CELL.Clear();
            for (int row_index = startRow; row_index <= endRow; row_index++)
            {
                for (int col_index = startCol; col_index <= endCol; col_index++)
                {
                    Object obj = ExcelAction.GetCellValue(template_worksheet, row_index, col_index);
                    if (obj != null)
                    {
                        String to_check = obj.ToString();
                        if (to_check.Contains(Variable_KEEP))
                        {
                            KEEP_ROW.Add(row_index);
                            KEEP_COL.Add(col_index);
                        }
                    }
                }
            }
            b_ret = true;
            return b_ret;
        }

        static public Boolean ReadKEEPCellContent(Worksheet report_worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            Boolean b_ret = false;
            KEEP_CELL.Clear();
            int count = KEEP_ROW.Count();
            while (count-- > 0)
            {
                int row = KEEP_ROW[count], col = KEEP_COL[count];
                Object obj = ExcelAction.GetCellValue(report_worksheet, row, col);
                KEEP_CELL.Add(obj);
            }
            b_ret = true;
            return b_ret;
        }
    }

    class ReportTemplate
    {
    }

    public class ReportContentPair
    {
        public String VariableName;
        public String VariableContent;
        public int VariableRow;
        public int VariableCol;
        public String LabelName;
        public int LabelRow;
        public int LabelCol;
        public Boolean On;
        public Boolean PreProcessing;
        public Boolean PostProcessing;

        public ReportContentPair(String labelName, String variableName)
        {
            InitialSetup(labelName, variableName);
        }
        public Boolean InitialSetup(String labelName, String variableName)
        {
            LabelName = labelName;
            VariableName = variableName;
            return !(String.IsNullOrWhiteSpace(LabelName) && String.IsNullOrWhiteSpace(VariableName));
        }
        public Boolean ToLabel(String label)
        {
            LabelName = label;
            return true;
        }
        public int Compare(String label)
        {
            return String.Compare(LabelName, label);
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
        static private int templateStartRow = 1;
        static private int templateStartCol = ExcelAction.ColumnNameToNumber("A");
        static private int templateEndRow = 22;
        static private int templateEndCol = ExcelAction.ColumnNameToNumber("N");
        static private List<ReportContentPair> templateContentPair = new List<ReportContentPair>();
        static public String[] DefaultLabel = {
            "Model Name", "Panel Module" , "AD Board", "Smart BD / OS Version", "Speaker / AQ Version", "Test Stage", "Judgement", "Sample S/N", "Tested by",
	        "Part No.", "T-Con Board", "Power Board", "Touch Sensor", "SW / PQ Version", "Test Period", "Approved by", 
            "Purpose:", "Condition:", "Equipment:", "Method:", "Criteria:", "Conclusion:", "Sample S/N:", "Test Period Start", "Test Period End",
            // dummy label for label-less
            "TITLE" };
        static public String[] DefaultVariable = {
            "ModelName", "PanelModule" , "ADBoard", "SmartBD_OSVersion", "Speaker_AQVersion", "TestStage", "Judgement", "SampleSN", "TestedBy",
	        "PartNo", "TConBoard", "PowerBoard", "TouchSensor", "SW_PQVersion", "TestPeriod", "Approvedby", "TestPeriodStart", "TestPeriodEnd",
            "Purpose", "Condition", "Equipment", "Method", "Criteria", "Conclusion", "SampleSN",
            // Some variable are label-less
            "Title" };

        static private Workbook wb_input_excel;
        static private Worksheet ws_input_excel;
        static private Worksheet ws_source_template;
        static private Worksheet ws_destination_template;
        static private String input_excel_filename;
        static List<ReportContentPair> source_rcp = new List<ReportContentPair>();
        static List<ReportContentPair> destination_rcp = new List<ReportContentPair>();

        static public List<ReportContentPair> SetupHeaderTemplateContentPair(Worksheet worksheet, List<String> labelList, List<String> variableList, Boolean AllowRepeatedVariable = false)
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

                    String labelName = label_list_to_search[index];
                    if (AllowRepeatedVariable == false)
                    {
                        variable_to_search.RemoveAt(index);
                        label_list_to_search.RemoveAt(index);
                    }

                    ReportContentPair cp = new ReportContentPair(labelName, variableName);
                    cp.VariableRow = row;
                    cp.VariableCol = col;
                    contentPairList.Add(cp);
                }
            }
            // search through excel until end of header section
            return contentPairList;
        }
        static public Boolean ProcessInputExcelAndTemplate(String input_excel_file)
        {
            // Open Source Header Template Excel workbook
            wb_input_excel = ExcelAction.OpenExcelWorkbook(filename: input_excel_file, ReadOnly: true);
            if (wb_input_excel == null)
            {
                LogMessage.WriteLine("ERR: Open workbook failed in ReadHeaderVariablesAccordingToTemplate(): " + input_excel_file);
                return false;
            }

            // Find source template sheet
            String source_template_sheet = InputExcel.SheetName_HeaderTemplate_Source;
            if (ExcelAction.WorksheetExist(wb_input_excel, source_template_sheet) == false)
            {
                LogMessage.WriteLine("ERR: source template worksheet doesn't exist on excel: " + input_excel_file);
                return false;
            }
            ws_source_template = ExcelAction.Find_Worksheet(wb_input_excel, source_template_sheet);

            // Find destination template sheet
            String destination_template_sheet = InputExcel.SheetName_HeaderTemplate_Destination;
            if (ExcelAction.WorksheetExist(wb_input_excel, destination_template_sheet) == false)
            {
                LogMessage.WriteLine("ERR: destination template worksheet doesn't exist on excel: " + input_excel_file);
                return false;
            }
            ws_destination_template = ExcelAction.Find_Worksheet(wb_input_excel, destination_template_sheet);

            String reportList_sheet = InputExcel.SheetName_ReportList;
            if (ExcelAction.WorksheetExist(wb_input_excel, reportList_sheet) == false)
            {
                LogMessage.WriteLine("ERR: Report List worksheet doesn't exist on excel: " + input_excel_file);
                return false;
            }
            ws_input_excel = ExcelAction.Find_Worksheet(wb_input_excel, InputExcel.SheetName_ReportList);

            input_excel_filename = input_excel_file;

            return true;
        }
        static public Boolean CreateHeaderVariablesLocationInfo(Worksheet sourceTemplateSheet, Worksheet destinationTemplateSheet)
        {
            source_rcp = SetupHeaderTemplateContentPair(sourceTemplateSheet, DefaultLabel.ToList(), DefaultVariable.ToList(), AllowRepeatedVariable: true);
            destination_rcp = SetupHeaderTemplateContentPair(destinationTemplateSheet, DefaultLabel.ToList(), DefaultVariable.ToList(), AllowRepeatedVariable: false);
            return true;
        }
        static public Boolean GetSourceAndDestinationVariableContent(Worksheet worksheet)
        {
            Boolean b_ret = false;

            foreach (ReportContentPair rcp in source_rcp)
            {
                String str = ExcelAction.GetCellTrimmedString(worksheet, rcp.VariableRow, rcp.VariableCol);
                rcp.VariableContent = str;
                foreach (ReportContentPair dest_rcp in destination_rcp)
                {
                    if (dest_rcp.VariableName == rcp.VariableName)
                    {
                        dest_rcp.VariableContent = str;
                        break;
                    }
                }
            }
            b_ret = true;
            return b_ret;
        }
        static public Boolean OutputDestinationVariableContent(Worksheet worksheet)
        {
            Boolean b_ret = false;

            foreach (ReportContentPair dest_rcp in destination_rcp)
            {
                ExcelAction.SetCellString(worksheet, dest_rcp.VariableRow, dest_rcp.VariableCol, dest_rcp.VariableContent);
            }
            b_ret = true;
            return b_ret;
        }
        static public Boolean ReplaceHeader(CopyReport report_to_copy)
        {
            Boolean b_ret = false;
            Workbook wb_source;
            Worksheet ws_source;

            String source_file = report_to_copy.Get_SRC_FullFilePath();
            String destination_file = report_to_copy.Get_DEST_FullFilePath();

            if (Storage.IsReportFilename(destination_file) == false)
            {
                // Do nothing if new filename does not look like a report filename.
                return false;
            }

            if (TestReport.OpenReportWorksheet(source_file, out wb_source, out ws_source) == false)
            {
                // Do nothing if opening excel & finding worksheet failed
                return false;
            }

            String destination_report_title = ReportGenerator.GetReportTitleAccordingToFilename(destination_file);
            String destination_report_sheetname = ReportGenerator.GetSheetNameAccordingToFilename(destination_file);
            String assignee = report_to_copy.DestinationAssignee;
            String today = DateTime.Now.ToString("yyyy/MM/dd");
            HeaderTemplate.ResetVariables();
            HeaderTemplate.UpdateVariables_TodayAssignee(today, assignee);
            HeaderTemplate.UpdateVariables_FilenameSheetname(filename: destination_report_title, sheetname: destination_report_sheetname);

            if (HeaderTemplate.ReadKEEPCellContent(ws_source, templateStartRow, templateStartCol, templateEndRow, templateEndCol) == false)
            {
                return false;
            }

            if (GetSourceAndDestinationVariableContent(ws_source) == false)
            {
                return false;
            }

            // paste new template
            ExcelAction.CopyPasteRows(ws_destination_template, ws_source, templateStartRow, templateEndRow);

            if (OutputDestinationVariableContent(ws_source) == false)
            {
                return false;
            }

            if (HeaderTemplate.PasteHeaderVariableContent(ws_source) == false)
            {
                return false;
            }

            if (HeaderTemplate.PasteKEEPCell(ws_source) == false)
            {
                return false;
            }

            // Something has been updated or always save (ex: to copy file & update) ==> save to excel file
            String destination_dir = Storage.GetDirectoryName(destination_file);
            // if parent directory does not exist, create recursively all parents
            if (Storage.DirectoryExists(destination_dir) == false)
            {
                Storage.CreateDirectory(destination_dir, auto_parent_dir: true);
            }
            ExcelAction.SaveExcelWorkbook(wb_source, filename: destination_file);

            ExcelAction.CloseExcelWorkbook(wb_source);

            b_ret = true;
            return b_ret;
        }
        static public Boolean UpdateTestReportHeader(out List<String> output_report_list, out String return_destination_path)
        {
            output_report_list = new List<String>();
            return_destination_path = "";

            Boolean bStillReadingExcel = true;
            // check title row
            const int row_start_index = 1;
            int row_index = row_start_index, col_index = 1;
            // start data processing since 2nd row
            row_index++;
            col_index = 1;
            List<CopyReport> report_list_to_be_processed = new List<CopyReport>();
            List<CopyReport> source_inexist_list = new List<CopyReport>();
            do
            {
                CopyReport ctp = new CopyReport();

                bStillReadingExcel = ctp.ReadFromExcelRow(ws_input_excel, row_index, col_index);
                if (bStillReadingExcel)
                {
                    // Because copy-only doesn't need to check report filename condition, such check is done later not here
                    // Here only checking whether source_file is available.
                    String source_filename = ctp.Get_SRC_FullFilePath();
                    if (Storage.FileExists(source_filename))
                    {
                        report_list_to_be_processed.Add(ctp);
                    }
                    else
                    {
                        source_inexist_list.Add(ctp);
                    }
                    row_index++;
                    col_index = 1;
                }
            }
            while (bStillReadingExcel);

            List<CopyReport> destination_not_report_filename_list = new List<CopyReport>();
            List<CopyReport> process_success_list = new List<CopyReport>();
            List<CopyReport> process_fail_list = new List<CopyReport>();

            // if valid file-list, sort it (when required) before further processing
            if (report_list_to_be_processed.Count > 0)
            {
                foreach (CopyReport cr in report_list_to_be_processed)
                {
                    Boolean success = false;

                    // only process when destination filename pass report/filename condition
                    if (Storage.IsReportFilename(cr.Get_DEST_FullFilePath()))
                    {
                        // To Be update;
                        success = ReplaceHeader(cr);

                        if (success)
                        {
                            process_success_list.Add(cr);
                        }
                        else
                        {
                            process_fail_list.Add(cr);
                        }
                    }
                    else
                    {
                        destination_not_report_filename_list.Add(cr);
                    }
                }
            }

            Boolean b_ret = true;
            if ((process_fail_list.Count > 0) || (destination_not_report_filename_list.Count > 0) || (source_inexist_list.Count > 0))
            {
                CopyReport.WriteErrorLog(wb_input_excel, ws_input_excel, source_inexist_list, destination_not_report_filename_list, process_fail_list);
                b_ret = false;
                string new_filename = Storage.GenerateFilenameWithDateTime(input_excel_filename);
                ExcelAction.CloseExcelWorkbook(workbook: wb_input_excel, SaveChanges: true, AsFilename: new_filename);
            }
            else
            {
                ExcelAction.CloseExcelWorkbook(wb_input_excel);
                b_ret = true;
            }

            return b_ret;
        }

        static public Boolean ChangeReportHeaderTemplateTask(String input_excel_file)
        {
            String return_destination_path;
            List<String> output_report_list;

            if (ProcessInputExcelAndTemplate(input_excel_file) == false)
                return false;
            if (CreateHeaderVariablesLocationInfo(ws_source_template, ws_destination_template) == false)
                return false;
            if (HeaderTemplate.FindKEEPCellLocation(ws_source_template, templateStartRow, templateStartCol, templateEndRow, templateEndCol) == false)
                return false;
            // Previous Header Variable are specified in destination template (because they are directly applied to destination report and no need to get content from source report
            if (HeaderTemplate.FindHeaderVariableLocation(ws_destination_template, templateStartRow, templateStartCol, templateEndRow, templateEndCol) == false)
                return false;
            if (UpdateTestReportHeader(out output_report_list, out return_destination_path) == false)
                return false;

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
