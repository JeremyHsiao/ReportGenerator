using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelReportApplication
{
    public enum ManPowerMemberIndex
    {
        Hierarchy = 0,
        Title,
        Project,
        Releases,
        Team,
        Assignee,
        Sprint,
        Target_start_date,
        Target_end_date,
        Due_date,
        Estimates,
        Parent,
        Priority,
        Labels,
        Components,
        Issue_key,
        Issue_status,
        Progress,
        Progress_completed,
        Progress_remaining,
        Progress_issue_count_IC,
        To_do_IC,
        In_progress_IC,
        Done_IC,
        Total_IC,
        Product_Type,       // for D0
        Man_hour,
        Customer,           // for D0
        COUNT,
    };

    public class ManPower
    {
        public String Hierarchy;
        public String Title;
        public String Project;
        public String Releases;
        public String Team;
        public String Assignee;
        public String Sprint;
        public String Target_start_date;
        public String Target_end_date;
        public String Due_date;
        public String Estimates;
        public String Parent;
        public String Priority;
        public String Labels;
        public String Components;
        public String Issue_key;
        public String Issue_status;
        public String Progress;
        public String Progress_completed;
        public String Progress_remaining;
        public String Progress_issue_count_IC;
        public String To_do_IC;
        public String In_progress_IC;
        public String Done_IC;
        public String Total_IC;
        public String Product_Type; // for D0
        public String Man_hour;
        public String Customer;     // for D0

        // generated data for each "Manpower" task (hierachy_string == Manpower)
        public String Task_Project_Name;
        public String Task_Action_Name;
        public String Task_Owner_Name;
        public DateTime Task_Start_Date; // Generate according to Target_start_date & Target_end_date;
        public DateTime Task_End_Date;
        public DateTime Task_Due_Date;
        public int Task_Start_Week;
        public int Task_End_Week;
        public Double ManHour;
        public String Daily_Average_ManHour_string;
        public Double Daily_Average_Manhour_value;
        // public String Daily_ManHour_String;      // Not used in current implementation
        public String Project_Action_Owner_WeekOfYear_ManHour;
        public List<String> Category_List;
        static public int Max_Category_Count;

        public ManPower ShallowCopy()
        {
            return (ManPower)this.MemberwiseClone();
        }

        public ManPower DeepCopy()
        {
            ManPower other = (ManPower)this.MemberwiseClone();
            return other;
        }

        // global data
        static public String Caption_Line;              // reading from CSV
        static public DateTime Start_Date, End_Date;    // search all CSV
        static public int Start_Week, End_Week;         // search all CSV
        static public List<Boolean> IsWorkingDay = new List<Boolean>();
        //static public String Title_StartDate_to_EndDate;  // Generated according to Start_Date, End_Date
        static public String Title_StartWeek_to_EndWeek;  // Generated according to Start_Date, End_Date
        static public String Title_Project_Action_Owner_WeekOfYear_ManHour;
        static public Dictionary<int, int> WorkingDayInWeek = new Dictionary<int, int>();
        static public List<int> WeekOfYearList = new List<int>();

        static public String hierarchy_string_for_project_v1 = "Task";
        static public String hierarchy_string_for_action_v1 = "Manpower";
        static public String hierarchy_string_for_project_v2 = "Manpower";
        static public String hierarchy_string_for_action_v2 = "Sub-Manpower";
        static public String hierarchy_string_for_D0_project = "Project";
        static public String hierarchy_string_for_D0_action = "Sub-Project";
        static public String empty_average_manhour = " ";
        static public Double empty_average_manhour_value = -1.0;

        static public Boolean hierarchy_D0_v1_detected = false;
        static public Boolean hierarchy_non_D0_v1_detected = false;
        static public Boolean hierarchy_auto_detected_finished = false;
        static public Boolean hierarchy_auto_detected_failed = false;

        static public String Recent_Task_Project_Name;

        // rounding digit for storing into variable "average_rounding_digit" after division calculation.
        static public int average_rounding_digit = 3;
        // rounding digit for storing into CSV data after weekly working-day * 
        static public String pSpecifier = "F1";   // floating-point with one digit after decimal


        public ManPower() { this.SetMemberByString(new List<String>()); }

        public ManPower(List<String> elements)
        {
            this.SetMemberByString(elements);
        }

        public Boolean Check_If_Hierarchy_Project()
        {
            if (hierarchy_auto_detected_failed)
                return false;

            if (hierarchy_D0_v1_detected)
            {
                // (1) hierarchy for D0 is (Project/Sub-Project)
                if (Hierarchy == hierarchy_string_for_D0_project)
                {
                    return true;
                }
            }
            else if (hierarchy_non_D0_v1_detected)
            {
                // (2) hierarchy non-D0 v1 is (Task/Manpower)
                if (Hierarchy == hierarchy_string_for_project_v1)
                {
                    return true;
                }
            }
            else
            {
                // (3) hierarchy non-D0 v2 is (Manpower/Sub-Manpower)
                if (Hierarchy == hierarchy_string_for_project_v2)
                {
                    return true;
                }
            }
            return false;
        }

        public Boolean Check_If_Hierarchy_Action()
        {
            if (hierarchy_auto_detected_failed)
                return false;

            if (hierarchy_D0_v1_detected)
            {
                // (1) hierarchy for D0 is (Project/Sub-Project)
                if (Hierarchy == hierarchy_string_for_D0_action)
                {
                    return true;
                }
            }
            else if (hierarchy_non_D0_v1_detected)
            {
                // (2) hierarchy non-D0 v1 is (Task/Manpower)
                if (Hierarchy == hierarchy_string_for_action_v1)
                {
                    return true;
                }
            }
            else
            {
                // (3) hierarchy non-D0 v1 is (Manpower/Sub-Manpower)
                if (Hierarchy == hierarchy_string_for_action_v2)
                {
                    return true;
                }
            }
            return false;
        }

        public void SetMemberByString(List<String> members)
        {
            int index_count = members.Count();
            if (index_count <= (int)ManPowerMemberIndex.COUNT)
            {
                String[] empty_str = new String[(int)ManPowerMemberIndex.COUNT - index_count];
                members.AddRange(empty_str);
            }

            int index = 0;
            Hierarchy = members[index++];
            Title = members[index++];
            Project = members[index++];
            Releases = members[index++];
            Team = members[index++];
            Assignee = members[index++];
            Sprint = members[index++];
            Target_start_date = members[index++];
            Target_end_date = members[index++];
            Due_date = members[index++];
            Estimates = members[index++];
            Parent = members[index++];
            Priority = members[index++];
            Labels = members[index++];
            Components = members[index++];
            Issue_key = members[index++];
            Issue_status = members[index++];
            Progress = members[index++];
            Progress_completed = members[index++];
            Progress_remaining = members[index++];
            Progress_issue_count_IC = members[index++];
            To_do_IC = members[index++];
            In_progress_IC = members[index++];
            Done_IC = members[index++];
            Total_IC = members[index++];

            // auto-detecting hierarchy for v1 (Task/Manpower)
            // v2 is (Manpower/Sub-Manpower)
            if (hierarchy_auto_detected_finished == false)
            {
                if (Hierarchy == hierarchy_string_for_project_v1)
                {
                    hierarchy_non_D0_v1_detected = true;
                    hierarchy_D0_v1_detected = false;
                    hierarchy_auto_detected_finished = true;
                    hierarchy_auto_detected_failed = false;
                }
                else if (Hierarchy == hierarchy_string_for_project_v2)
                {
                    hierarchy_non_D0_v1_detected = false;
                    hierarchy_D0_v1_detected = false;
                    hierarchy_auto_detected_finished = true;
                    hierarchy_auto_detected_failed = false;
                }
                else if (Hierarchy == hierarchy_string_for_D0_project)
                {
                    hierarchy_D0_v1_detected = true;
                    hierarchy_auto_detected_finished = true;
                    hierarchy_auto_detected_failed = false;
                }
                else
                {
                    hierarchy_D0_v1_detected = false;
                    hierarchy_auto_detected_finished = true;
                    hierarchy_auto_detected_failed = true;
                }
            }

            // For D0
            if (hierarchy_D0_v1_detected)
            {
                Product_Type = members[index++];
                Man_hour = members[index++];
                Customer = members[index++];
            }
            else
            {
                Man_hour = members[index++];
            }

            // Post-processing
            if (Check_If_Hierarchy_Project())
            {
                ManPower.Recent_Task_Project_Name = Title;
            }
            else if (Check_If_Hierarchy_Action())  // only man-power to calculate average man-hour
            {
                Process_ManPower_Data();
            }
            else
            {
            }

            //if (hierarchy_D0_v1_detected)
            //{
            //    // (1) hierarchy for D0 is (Project/Sub-Project)
            //    if (Hierarchy == hierarchy_string_for_D0_project)
            //    {
            //        ManPower.Recent_Task_Project_Name = Title;
            //    }
            //    else if (Hierarchy == hierarchy_string_for_D0_action)  // only man-power to calculate average man-hour
            //    {
            //        Process_ManPower_Data();
            //    }
            //    else
            //    {
            //    }
            //}
            //else if (hierarchy_non_D0_v1_detected)
            //{
            //    // (1) hierarchy non-D0 v1 is (Task/Manpower)
            //    if (Hierarchy == hierarchy_string_for_project_v1)
            //    {
            //        ManPower.Recent_Task_Project_Name = Title;
            //    }
            //    else if (Hierarchy == hierarchy_string_for_action_v1)  // only man-power to calculate average man-hour
            //    {
            //        Process_ManPower_Data();
            //    }
            //    else
            //    {
            //    }
            //}
            //else
            //{
            //    // (1) hierarchy non-D0 v1 is (Manpower/Sub-Manpower)
            //    if (Hierarchy == hierarchy_string_for_project_v2)
            //    {
            //        ManPower.Recent_Task_Project_Name = Title;
            //    }
            //    else if (Hierarchy == hierarchy_string_for_action_v2)  // only man-power to calculate average man-hour
            //    {
            //        Process_ManPower_Data();
            //    }
            //    else
            //    {
            //    }
            //}
        }
        //// generated data for each "Manpower" task (hierachy_string == Manpower)
        //public String Task_Project_Name;
        //public String Task_Item_Name;
        //public String Task_Assignee_Name;

        public void Process_ManPower_Data()
        {
            Boolean average_can_be_calculated = true;

            Task_Project_Name = ManPower.Recent_Task_Project_Name;
            Task_Action_Name = Title;
            Task_Owner_Name = Assignee;

            if (String.IsNullOrWhiteSpace(Target_start_date))
            {
                average_can_be_calculated = false;
            }
            else
            {
                Task_Start_Date = Convert.ToDateTime(Target_start_date, DateOnly.datetime_culture).Date;
                Task_Start_Week = YearWeek.GetYearAndWeekOfYear(Task_Start_Date);
            }

            if (String.IsNullOrWhiteSpace(Target_end_date))
            {
                average_can_be_calculated = false;
            }
            else
            {
                Task_End_Date = Convert.ToDateTime(Target_end_date, DateOnly.datetime_culture).Date;
                Task_End_Week = YearWeek.GetYearAndWeekOfYear(Task_End_Date);
            }

            if (Double.TryParse(Man_hour, out ManHour) == false)
            {
                average_can_be_calculated = false;
                ManHour = 0;
            }
            else
            {
                // ManHour is valid.
            }

            // Start date shouldn't be later than End date
            if (Task_End_Date < Task_Start_Date)
            {
                average_can_be_calculated = false;
            }

            if (average_can_be_calculated)
            {
                DateTime start = Task_Start_Date;
                DateTime end = Task_End_Date;
                int workday_count = DateOnly.BusinessDaysUntil(start, end);
                // workday must be > 0 (ie, from start to end date shouldn't be all in the middle of holidays)
                if (workday_count > 0)
                {
                    Double average_man_hour = Math.Round(ManHour / workday_count, ManPower.average_rounding_digit);
                    Daily_Average_ManHour_string = average_man_hour.ToString(ManPower.pSpecifier);
                    Daily_Average_Manhour_value = average_man_hour;
                }
                else
                {
                    // man-power plan needs to be checked and updated.
                    Daily_Average_ManHour_string = empty_average_manhour;
                    Daily_Average_Manhour_value = empty_average_manhour_value;
                }
            }
            else
            {
                // man-power plan needs to be checked and updated.
                Daily_Average_ManHour_string = empty_average_manhour;
                Daily_Average_Manhour_value = empty_average_manhour_value;
            }
        }

        static public String AddComma(String item)
        {
            String return_string = item + ',';
            return return_string;
        }

        static public String AddQuote(String item)
        {
            // For csv file output, double-quotation must be repreated once.
            item = item.Replace("\"", "\"\"");
            String return_string = '"' + item + '"';
            return return_string;
        }

        static public String AddQuoteWithComma(String item)
        {
            String return_string = AddComma(AddQuote(item));
            return return_string;
        }

        // this function is static
        static public String GenerateDateTitle(DateTime start, DateTime end)
        {
            String ret_str = "";
            if (start > end)
            {
                // to-check: shouldn't be here
            }
            else
            {
                // At least one date (start_date)
                DateTime dt = start.Date;
                ret_str = dt.ToString("d", DateOnly.datetime_culture);
                dt = dt.AddDays(1.0);
                // add "," + next-date till next-date is the end-date
                while (dt <= end.Date)
                {
                    ret_str += "," + dt.ToString("d", DateOnly.datetime_culture);
                    dt = dt.AddDays(1.0);
                }
                // reaching here when the next-date is after the end-date
            }
            return ret_str;
        }

        static public String GenerateWeekOfYearTitle(DateTime start, DateTime end)
        {
            String ret_str = "";
            if (start > end)
            {
                // to-check: shouldn't be here
            }
            else
            {
                // At least one date (start_date)
                DateTime dt = start.Date;

                ret_str = dt.ToString("yyyy", DateOnly.datetime_culture).Substring(3, 1) + YearWeek.GetYearAndWeekOfYear(dt).ToString();
                dt = dt.AddDays(7.0);
                // add "," + next-date till next-date is the end-date
                while (dt <= end.Date)
                {
                    ret_str += "," + dt.ToString("yyyy", DateOnly.datetime_culture).Substring(3, 1) + YearWeek.GetYearAndWeekOfYear(dt).ToString();
                    dt = dt.AddDays(7.0);
                }
                // reaching here when the next-date is after the end-date
            }
            return ret_str;
        }

        // this function is working properly when title start/end date are set up correctly.
        //public String GenerateManPowerDailyEffortString()
        //{
        //    String ret_str = "";

        //    // check if (1) a man-power item (2) Average ManHour is not empty (3) start/end date is not correct
        //    if ((this.Check_If_Hierarchy_Action()) || (this.Daily_Average_ManHour == empty_average_manhour) ||
        //        (this.Task_Start_Date > this.Task_End_Date) || (ManPower.Start_Date > ManPower.End_Date))
        //    {
        //        // to-check: shouldn't be here
        //        return ret_str;
        //    }

        //    // Find overlay with Task_Start/Task_End -- by default
        //    DateTime overlay_start = ManPower.Start_Date, overlay_end = ManPower.End_Date;

        //    // check 1: Task start is later than Man Power Start or not? later one will be the new overlay start_date
        //    if (this.Task_Start_Date > overlay_start)
        //    {
        //        overlay_start = this.Task_Start_Date;
        //    }

        //    // check 2: Task end date is earlier than Man Power End date or not? earlier one will be the new overlay end_date
        //    if (this.Task_End_Date < overlay_end)
        //    {
        //        overlay_end = this.Task_End_Date;
        //    }

        //    int overlay_start_index = (int)(overlay_start - ManPower.Start_Date).TotalDays,
        //        overlay_end_index = (int)(overlay_end - ManPower.Start_Date).TotalDays,
        //        total_end_index = (int)(ManPower.End_Date - ManPower.Start_Date).TotalDays;

        //    // 1st day is already overlay-date or not? if yes, average-manhour for working day or "0" for holiday
        //    // if 1st day is not-yet an overlay-date, fill "0"
        //    if (overlay_start_index == 0)
        //    {
        //        ret_str += (ManPower.IsWorkingDay[0]) ? this.Daily_Average_ManHour : "0";
        //    }
        //    else
        //    {
        //        ret_str += "0";
        //    }

        //    int date_index = 1;

        //    // before overlay
        //    while (date_index < overlay_start_index)
        //    {
        //        ret_str += ", 0";
        //        date_index++;
        //    }

        //    // during overlay -- output average man-hour
        //    while (date_index <= overlay_end_index)
        //    {
        //        ret_str += ", ";
        //        ret_str += (ManPower.IsWorkingDay[date_index]) ? this.Daily_Average_ManHour : "0";
        //        date_index++;
        //    }

        //    // after overlay
        //    while (date_index <= total_end_index)
        //    {
        //        ret_str += ", 0";
        //        date_index++;
        //    }


        //    //DateTime filling_Date = ManPower.Start_Date;

        //    //// 1st date is already overlay-date?
        //    //if (filling_Date == overlay_start)
        //    //{
        //    //    ret_str += this.Average_ManHour;    // if yes, output average manhour
        //    //}
        //    //else
        //    //{
        //    //    ret_str += "0";
        //    //}
        //    //filling_Date = filling_Date.AddDays(1.0);

        //    //// before overlay
        //    //while(filling_Date < overlay_start) 
        //    //{
        //    //    ret_str += ", 0";
        //    //    filling_Date = filling_Date.AddDays(1.0);
        //    //}

        //    //// during overlay -- output average man-hour
        //    //while(filling_Date <= overlay_end)
        //    //{
        //    //    ret_str += ", ";
        //    //    if (DateOnly.IsHoliday(filling_Date))
        //    //    {
        //    //        ret_str += "0";
        //    //    }
        //    //    else
        //    //    {
        //    //        ret_str += this.Average_ManHour; 
        //    //    }
        //    //    filling_Date = filling_Date.AddDays(1.0);
        //    //}

        //    //// after overlay
        //    //while (filling_Date <= ManPower.End_Date)
        //    //{
        //    //    ret_str += ", 0";
        //    //    filling_Date = filling_Date.AddDays(1.0);
        //    //}

        //    this.Daily_ManHour_String = ret_str;
        //    return ret_str;
        //}

        public override String ToString()
        {
            String return_string;

            return_string = AddQuoteWithComma(this.Hierarchy);
            return_string += AddQuoteWithComma(this.Title);
            return_string += AddQuoteWithComma(this.Project);
            return_string += AddQuoteWithComma(this.Releases);
            return_string += AddQuoteWithComma(this.Team);
            return_string += AddQuoteWithComma(this.Assignee);
            return_string += AddQuoteWithComma(this.Sprint);
            return_string += AddQuoteWithComma(this.Target_start_date);
            return_string += AddQuoteWithComma(this.Target_end_date);
            return_string += AddQuoteWithComma(this.Due_date);
            return_string += AddQuoteWithComma(this.Estimates);
            return_string += AddQuoteWithComma(this.Parent);
            return_string += AddQuoteWithComma(this.Priority);
            return_string += AddQuoteWithComma(this.Labels);
            return_string += AddQuoteWithComma(this.Components);
            return_string += AddQuoteWithComma(this.Issue_key);
            return_string += AddQuoteWithComma(this.Issue_status);
            return_string += AddComma(this.Progress);
            return_string += AddComma(this.Progress_completed);
            return_string += AddComma(this.Progress_remaining);
            return_string += AddComma(this.Progress_issue_count_IC);
            return_string += AddComma(this.To_do_IC);
            return_string += AddComma(this.In_progress_IC);
            return_string += AddComma(this.Done_IC);
            return_string += AddComma(this.Total_IC);
            // For D0
            if ((Hierarchy == hierarchy_string_for_D0_project) || (Hierarchy == hierarchy_string_for_D0_action))
            {
                return_string += AddQuoteWithComma(this.Product_Type);
                return_string += AddComma(this.Man_hour);
                return_string += AddQuote(this.Customer);
            }
            else
            {
                return_string += this.Man_hour;  // no need to output comma
            }
            return return_string;
        }
    }

    static public class ManPowerTask
    {

        //static public void ReadManPowerTaskCSV(String csv_filename)
        //{
        //    Excel.Workbook wb;
        //    String new_filename = Storage.GenerateFilenameWithDateTime(Filename: csv_filename, FileExt: ".xlsx");
        //    wb = ExcelAction.OpenCSV(csv_filename);
        //    ExcelAction.CloseCSV_SaveAsExcel(workbook: wb, SaveChanges: true, AsFilename: new_filename);
        //}

        static public List<ManPower> ReadManPowerTaskCSV(String csv_filename)
        {
            DateOnly.SortHoliday();

            List<ManPower> ret_manpower_list = new List<ManPower>();
            using (TextFieldParser csvParser = new TextFieldParser(csv_filename))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                // Skip the row with the column names
                ManPower.Caption_Line = csvParser.ReadLine();

                while (!csvParser.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    List<String> elements = new List<String>();
                    elements.AddRange(csvParser.ReadFields());
                    ManPower manpower = new ManPower(elements);
                    List<String> category_list = Extracting_Category(manpower.Title);
                    manpower.Category_List = category_list;
                    if (category_list.Count > ManPower.Max_Category_Count)
                    {
                        ManPower.Max_Category_Count = category_list.Count;
                    }
                    ret_manpower_list.Add(manpower);
                }
            }
            return ret_manpower_list;
        }

        //static public List<ManPower> Post_Processing(List<ManPower> list_before_post_processing)
        //{
        //    List<ManPower> ret_manpower_list = list_before_post_processing;

        //    // Generated data for ManPower
        //    ManPower.Start_Date = DateOnly.FindEearliestTargetStartDate(ret_manpower_list);
        //    ManPower.End_Date = DateOnly.FindLatestTargetEndDate(ret_manpower_list);
        //    DateOnly.Update_Holiday_Range(ManPower.Start_Date, ManPower.End_Date);
        //    ManPower.IsWorkingDay.Clear();
        //    for (DateTime dt = ManPower.Start_Date.Date; dt <= ManPower.End_Date.Date; dt = dt.AddDays(1.0))
        //    {
        //        if (DateOnly.IsHoliday(dt))
        //        {
        //            ManPower.IsWorkingDay.Add(false);       // a holiday --> not a working day
        //        }
        //        else
        //        {
        //            ManPower.IsWorkingDay.Add(true);
        //        }
        //    }
        //    ManPower.Title_StartDate_to_EndDate = ManPower.GenerateDateTitle(ManPower.Start_Date, ManPower.End_Date);
        //    ManPower.Title_StartWeek_to_EndWeek = ManPower.GenerateWeekOfYearTitle(ManPower.Start_Date, ManPower.End_Date);

        //    // Setup static class YearWeek variables
        //    YearWeek.SetupByStartDateEndDate(ManPower.Start_Date, ManPower.End_Date);
        //    ManPower.Start_Week = YearWeek.GetStartWeek();
        //    ManPower.End_Week = YearWeek.GetEndWeek();

        //    // Generated data for each task
        //    foreach (ManPower mp in ret_manpower_list)
        //    {
        //        mp.GenerateManPowerDailyEffortString();
        //    }

        //    return ret_manpower_list;
        //}

        static public List<ManPower> Processing_DateWeekHoliday(List<ManPower> list_before_post_processing)
        {
            // Generated data for ManPower
            ManPower.Start_Date = DateOnly.FindEearliestTargetStartDate(list_before_post_processing);
            ManPower.End_Date = DateOnly.FindLatestTargetEndDate(list_before_post_processing);
            DateOnly.Update_Holiday_Range(ManPower.Start_Date, ManPower.End_Date);
            ManPower.IsWorkingDay.Clear();
            for (DateTime dt = ManPower.Start_Date.Date; dt <= ManPower.End_Date.Date; dt = dt.AddDays(1.0))
            {
                if (DateOnly.IsHoliday(dt))
                {
                    ManPower.IsWorkingDay.Add(false);       // a holiday --> not a working day
                }
                else
                {
                    ManPower.IsWorkingDay.Add(true);
                }
            }
            //ManPower.Title_StartDate_to_EndDate = ManPower.GenerateDateTitle(ManPower.Start_Date, ManPower.End_Date);
            ManPower.Title_StartWeek_to_EndWeek = ManPower.GenerateWeekOfYearTitle(ManPower.Start_Date, ManPower.End_Date);

            // Setup static class YearWeek variables
            YearWeek.SetupByStartDateEndDate(ManPower.Start_Date, ManPower.End_Date);
            ManPower.Start_Week = YearWeek.GetStartWeek();
            ManPower.End_Week = YearWeek.GetEndWeek();

            return list_before_post_processing;
        }

        //static public Boolean OutputManPowerTaskCSV(String csv_output)
        //{
        //    Boolean ret = false;

        //    //before your loop
        //    var csv = new StringBuilder();

        //    //in your loop
        //    var first = reader[0].ToString();
        //    var second = image.ToString();
        //    //Suggestion made by KyleMit
        //    var newLine = string.Format("{0},{1}", first, second);
        //    csv.AppendLine(newLine);

        //    //after your loop
        //    File.WriteAllText(csv_output, csv.ToString());
        //    return ret;
        //}

        /*
        //// return: 
        //// (1) earliest & latest Target start date
        //// (2) earliest & latest Target end date
        //// (3) earliest & latest Target due date
        //static private List<DateTime> GatherDateInfo(List<ManPower> manpower)
        //{
        //    DateTime earliest_target_start_date = DateTime_Latest; // default when no earliest date found
        //    DateTime latest_target_start_date = DateTime_Earliest; // default when no latest date found

        //    foreach (ManPower mp in manpower)
        //    {
        //        // for target_start_Date
        //        DateTime target_start_date = Convert.ToDateTime(mp.Target_start_date);
        //        // find earliest
        //        if (target_start_date < earliest_dt)
        //        {
        //            earliest_dt = target_start_date;
        //        }
        //        // find latest
        //        if (target_start_date > latest_dt)
        //        {
        //            latest_dt = target_start_date;
        //        }

        //        // for target_end_Date
        //        DateTime target_end_date = Convert.ToDateTime(mp.Target_end_date);
        //        // find earliest
        //        if (target_end_date < earliest_dt)
        //        {
        //            earliest_dt = target_start_date;
        //        }
        //        // find latest
        //        if (target_start_date > latest_dt)
        //        {
        //            latest_dt = target_start_date;
        //        }

        //    }
        //    return earliest_dt;
        //}
        */

        //static public void ProcessManPowerPlan(String manpower_csv)
        //{
        //    List<ManPower> manpower_list_before = ReadManPowerTaskCSV(manpower_csv);
        //    List<ManPower> manpower_list = Post_Processing(manpower_list_before);
        //    //DateTime manpower_due_date = FindLatestDueDate(manpower_list);

        //    //before your loop
        //    var csv = new StringBuilder();

        //    //csv.AppendLine(ManPower.Caption_Line);
        //    csv.AppendLine(ManPower.Caption_Line + "," + ManPower.Title_StartDate_to_EndDate);

        //    //in your loop
        //    foreach (ManPower mp in manpower_list)
        //    {
        //        //var newLine = mp.ToString();
        //        var newLine = mp.ToString() + "," + mp.Daily_ManHour_String;
        //        csv.AppendLine(newLine);
        //    }

        //    //after your loop
        //    File.WriteAllText(Storage.GenerateFilenameWithDateTime(manpower_csv, ".csv"), csv.ToString(), Encoding.UTF8);
        //}


        static public List<String> Extracting_Category(String title)
        {
            List<String> ret_list = new List<string>();

            String pattern = @"(\[\w+\s*\w*\])";
            Regex rgx = new Regex(pattern);
            Match match = Regex.Match(title, pattern, RegexOptions.None);
            while (match.Success)
            {
                String category_str = match.Value;
                ret_list.Add(category_str);
                match = match.NextMatch();
            }
            return ret_list;
        }

        static public void ProcessManPowerPlan_V2(String manpower_csv)
        {
            ManPower.hierarchy_auto_detected_finished = false;
            List<ManPower> manpower_list_before = ReadManPowerTaskCSV(manpower_csv);
            List<ManPower> manpower_list = Processing_DateWeekHoliday(manpower_list_before);
            //DateTime manpower_due_date = FindLatestDueDate(manpower_list);

            //before your loop
            var csv = new StringBuilder();

            // Setup Title line
            ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour = "ProjectStage,TestAction,Owner,Week,ManHourThisWeek,";
            //csv.AppendLine(ManPower.Caption_Line);
            csv.AppendLine(ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour + ManPower.Caption_Line);

            // Setup & repeat for weekly man-hour
            String Empty_Field_String = ",,,,,";

            // add items in week of year
            foreach (ManPower mp in manpower_list)
            {
                if (mp.Check_If_Hierarchy_Project())
                {
                    csv.AppendLine(Empty_Field_String + mp.ToString());
                }
                else if (mp.Check_If_Hierarchy_Action())   // only man-power to calculate average man-hour
                {
                    // need to deal with 1st week and last week of this task

                    String Item_Field_String = ManPower.AddQuoteWithComma(mp.Task_Project_Name);
                    Item_Field_String += ManPower.AddQuoteWithComma(mp.Task_Action_Name);
                    Item_Field_String += ManPower.AddQuoteWithComma(mp.Task_Owner_Name);
                    int begin_wk_index = YearWeek.IndexOf(mp.Task_Start_Week);
                    int end_wk_index = YearWeek.IndexOf(mp.Task_End_Week);
                    int current_wk_index = begin_wk_index;
                    Double daily_average_manhour_value = mp.Daily_Average_Manhour_value;
                    // for 1st week
                    if (current_wk_index == begin_wk_index)
                    {
                        DateTime first_date = mp.Task_Start_Date;
                        int remaining_workday = YearWeek.WeekWorkdayToLastDateFrom(first_date);
                        //special case: 1st week is also last week
                        if (current_wk_index == end_wk_index)
                        {
                            DateTime last_date = mp.Task_End_Date;
                            if (DateOnly.IsHoliday(last_date))
                            {
                                remaining_workday -= YearWeek.WeekWorkdayToLastDateFrom(last_date);
                            }
                            else
                            {
                                remaining_workday -= YearWeek.WeekWorkdayToLastDateFrom(last_date);
                                remaining_workday++;
                            }
                        }
                        Double weekly_manhour = remaining_workday * daily_average_manhour_value;
                        int current_yearweek = YearWeek.ElementAt(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + mp.ToString());
                        current_wk_index++;
                    }

                    // for 2nd until one-week before last week
                    while (current_wk_index <= end_wk_index - 1)
                    {
                        Double weekly_manhour = YearWeek.GetWorkingDayOfWeekByIndex(current_wk_index) * daily_average_manhour_value;
                        int current_yearweek = YearWeek.ElementAt(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + mp.ToString());
                        current_wk_index++;
                    }
                    // for last week

                    if (current_wk_index == end_wk_index)
                    {
                        DateTime last_date = mp.Task_End_Date;
                        int remaining_workday = YearWeek.WeekWorkdayFromFirstDateTo(last_date);
                        Double weekly_manhour = remaining_workday * daily_average_manhour_value;
                        int current_yearweek = YearWeek.ElementAt(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + mp.ToString());
                        current_wk_index++;
                    }
                }

            }

            //after your loop
            File.WriteAllText(Storage.GenerateFilenameWithDateTime(manpower_csv, ".csv"), csv.ToString(), Encoding.UTF8);
        }

        static public void ProcessManPowerPlan_V3(String manpower_csv)
        {
            ManPower.hierarchy_auto_detected_finished = false;
            List<ManPower> manpower_list_before = ReadManPowerTaskCSV(manpower_csv);
            List<ManPower> manpower_list = Processing_DateWeekHoliday(manpower_list_before);
            //DateTime manpower_due_date = FindLatestDueDate(manpower_list);

            //before your loop
            var csv = new StringBuilder();

            // Setup Title line
            // Setup & repeat for weekly man-hour
            String Empty_Field_String = ",,,,,";
            ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour = ManPower.AddQuoteWithComma("ProjectStage") +
                                                                        ManPower.AddQuoteWithComma("TestAction") +
                                                                        ManPower.AddQuoteWithComma("Owner") +
                                                                        ManPower.AddQuoteWithComma("Week") +
                                                                        ManPower.AddQuoteWithComma("ManHourThisWeek");
            for (int index = 0; index < ManPower.Max_Category_Count; index++)
            {
                ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddComma("Category" + index.ToString());
            }
            //csv.AppendLine(ManPower.Caption_Line);
            csv.AppendLine(ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour + ManPower.Caption_Line);

            // add items in week of year
            foreach (ManPower mp in manpower_list)
            {
                int category_countdown = ManPower.Max_Category_Count;
                String category_string = "";
                foreach (String str in mp.Category_List)
                {
                    category_string += ManPower.AddQuoteWithComma(str);
                    category_countdown--;
                }
                while (category_countdown-- > 0)
                {
                    category_string += ",";
                }

                if (mp.Check_If_Hierarchy_Project())
                {
                    csv.AppendLine(Empty_Field_String + category_string + mp.ToString());
                }
                else if (mp.Check_If_Hierarchy_Action())   // only man-power to calculate average man-hour
                {
                    // need to deal with 1st week and last week of this task

                    String Item_Field_String = ManPower.AddQuoteWithComma(mp.Task_Project_Name);
                    Item_Field_String += ManPower.AddQuoteWithComma(mp.Task_Action_Name);
                    Item_Field_String += ManPower.AddQuoteWithComma(mp.Task_Owner_Name);
                    int begin_wk_index = YearWeek.IndexOf(mp.Task_Start_Week);
                    int end_wk_index = YearWeek.IndexOf(mp.Task_End_Week);
                    int current_wk_index = begin_wk_index;
                    Double daily_average_manhour_value = mp.Daily_Average_Manhour_value;
                    // for 1st week
                    if (current_wk_index == begin_wk_index)
                    {
                        DateTime first_date = mp.Task_Start_Date;
                        int remaining_workday = YearWeek.WeekWorkdayToLastDateFrom(first_date);
                        //special case: 1st week is also last week
                        if (current_wk_index == end_wk_index)
                        {
                            DateTime last_date = mp.Task_End_Date;
                            if (DateOnly.IsHoliday(last_date))
                            {
                                remaining_workday -= YearWeek.WeekWorkdayToLastDateFrom(last_date);
                            }
                            else
                            {
                                remaining_workday -= YearWeek.WeekWorkdayToLastDateFrom(last_date);
                                remaining_workday++;
                            }
                        }
                        Double weekly_manhour = remaining_workday * daily_average_manhour_value;
                        int current_yearweek = YearWeek.ElementAt(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + category_string + mp.ToString());
                        current_wk_index++;
                    }

                    // for 2nd until one-week before last week
                    while (current_wk_index <= end_wk_index - 1)
                    {
                        Double weekly_manhour = YearWeek.GetWorkingDayOfWeekByIndex(current_wk_index) * daily_average_manhour_value;
                        int current_yearweek = YearWeek.ElementAt(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + category_string + mp.ToString());
                        current_wk_index++;
                    }
                    // for last week

                    if (current_wk_index == end_wk_index)
                    {
                        DateTime last_date = mp.Task_End_Date;
                        int remaining_workday = YearWeek.WeekWorkdayFromFirstDateTo(last_date);
                        Double weekly_manhour = remaining_workday * daily_average_manhour_value;
                        int current_yearweek = YearWeek.ElementAt(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + category_string + mp.ToString());
                        current_wk_index++;
                    }
                }

            }

            //after your loop
            File.WriteAllText(Storage.GenerateFilenameWithDateTime(manpower_csv, ".csv"), csv.ToString(), Encoding.UTF8);
        }

        static public List<ManPowerHolidayList> SetupHolidayListFromCSV(String csv_file)
        {
            // Create a list of site-based empty ManPowerHolidayList (holiday to be added later) 
            List<ManPowerHolidayList> ret_list_of_site_holidays = new List<ManPowerHolidayList>();
            List<Site> holiday_site_list = new List<Site>();

            foreach (String site_str in Site.SiteList)
            {
                // setup site info
                Site this_site = new Site();
                this_site.Name = site_str;
                holiday_site_list.Add(this_site);

                // Add holiday-list & associate site info
                ManPowerHolidayList holidays = new ManPowerHolidayList();
                holidays.Site = this_site;                                  // Associate site to this holiday list
                ret_list_of_site_holidays.Add(holidays);                    // Add empty site holiday
            }

            // parsing title to decide which site is for
            using (TextFieldParser csvParser = new TextFieldParser(csv_file))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = false;        // no quotation in current csv

                List<String> title = new List<String>();
                List<Site> title_site = new List<Site>();

                // Read Title row
                if (!csvParser.EndOfData)
                {
                    title.AddRange(csvParser.ReadFields());
                    foreach (String item in title)
                    {
                        Site this_site = new Site();
                        this_site.Name = item;
                        title_site.Add(this_site);
                    }
                }

                // Fill Holiday List according to title_site
                List<String> elements = new List<String>();
                while (!csvParser.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    elements.AddRange(csvParser.ReadFields());
                    int col_index = -1;
                    foreach (String item in elements)
                    {
                        col_index++;
                        if (String.IsNullOrWhiteSpace(item))
                            continue;

                        // get ManPowerDate
                        ManPowerDate mp_date = new ManPowerDate();
                        mp_date.FromString(item);
                        Site which_site = title_site[col_index];
                        ManPowerHolidayList site_holiday = ret_list_of_site_holidays[which_site.Index];
                        site_holiday.Add(mp_date);
                    }
                }
            }

            return ret_list_of_site_holidays;
        }

    }

    static public class YearWeek
    {
        static private DateTime StartDate;
        static private DateTime EndDate;
        static private List<int> yearweek_list = new List<int>();
        static private List<int> weekly_workday_list = new List<int>();
        static private Dictionary<DateTime, int> remaining_workday_to_Saturday_from = new Dictionary<DateTime, int>();
        static private Dictionary<DateTime, int> remaining_workday_from_Sunday_to = new Dictionary<DateTime, int>();
        static public int invalid_index = -1;
        static public int invalid_yearweek = -1;

        static public List<int> YearWeekList() { return yearweek_list; }

        static public List<int> WeeklyWorkdayList() { return weekly_workday_list; }

        static public int WeekWorkdayToLastDateFrom(DateTime datetime)
        {
            if (DateOnly.IsBetween(datetime, StartDate, EndDate))
            {
                return remaining_workday_to_Saturday_from[datetime];
            }
            else
            {
                return 0;
            }
        }

        static public int WeekWorkdayFromFirstDateTo(DateTime datetime)
        {
            if (DateOnly.IsBetween(datetime, StartDate, EndDate))
            {
                return remaining_workday_from_Sunday_to[datetime];
            }
            else
            {
                return 0;
            }
        }

        static public void SetupByStartDateEndDate(DateTime start, DateTime end)
        {
            YearWeek.StartDate = start;
            YearWeek.EndDate = end;
            yearweek_list.Clear();
            weekly_workday_list.Clear();
            remaining_workday_to_Saturday_from.Clear();
            remaining_workday_from_Sunday_to.Clear();

            int weekofyear_temp = GetYearAndWeekOfYear(start);
            // temporarily accumulating workday for all processed date within this week
            int ACC_workday_from_Sunday = 0;
            // temporarily store all processed date within this week
            List<DateTime> datetime_thisweek = new List<DateTime>();
            // temporarily store workday from Sunday/StartDate to this Date for all processed date within this week
            List<Boolean> is_weekly_workday = new List<Boolean>();
            for (DateTime dateime = start; dateime <= end; dateime = dateime.AddDays(1.0))
            {
                switch (dateime.DayOfWeek)
                {
                    case DayOfWeek.Sunday:

                        datetime_thisweek.Add(dateime);
                        // accumulator = 0 & record is_w;
                        ACC_workday_from_Sunday = 0;
                        is_weekly_workday.Add(false);

                        // update weekofyear here 
                        weekofyear_temp = GetYearAndWeekOfYear(dateime);

                        // record processed date
                        remaining_workday_from_Sunday_to.Add(dateime, ACC_workday_from_Sunday);

                        break;
                    case DayOfWeek.Saturday:

                        // datetime no need to add and will be cleared later
                        //datetime_thisweek.Add(dateime);
                        // accumulator don't change (holiday) & record remaining_workday_from_Sunday_to;
                        //workday_since_Sunday_acc++;

                        //workday_1st_day_of_week_till_datetime.Add(workday_since_Sunday_acc);
                        // store weekofyear/working_day info here
                        weekly_workday_list.Add(ACC_workday_from_Sunday);
                        yearweek_list.Add(weekofyear_temp);

                        // record workday_datetime_till_last_day_of_week
                        // updating by looping datetime_thisweek
                        int temp_workday = 0;
                        int workday_to_Saturday;
                        for (int index = 0; index < datetime_thisweek.Count; index++)
                        {
                            DateTime dt = datetime_thisweek[index];
                            if (is_weekly_workday[index])
                            {
                                workday_to_Saturday = ACC_workday_from_Sunday - temp_workday;
                                temp_workday++;
                            }
                            else
                            {
                                workday_to_Saturday = ACC_workday_from_Sunday - temp_workday;
                            }
                            remaining_workday_to_Saturday_from.Add(dt, workday_to_Saturday);
                        }
                        remaining_workday_to_Saturday_from.Add(dateime, 0);

                        remaining_workday_from_Sunday_to.Add(dateime, ACC_workday_from_Sunday); ; // no workday on Saturday
                        // clear after it is recorded
                        ACC_workday_from_Sunday = 0;
                        datetime_thisweek.Clear();
                        is_weekly_workday.Clear();

                        break;
                    default:
                        // the same current_weekofyear
                        datetime_thisweek.Add(dateime);
                        // accumulator++ for working day & record remaining_workday_from_Sunday_to;
                        if (DateOnly.IsHoliday(dateime))
                        {
                            is_weekly_workday.Add(false);
                        }
                        else
                        {
                            is_weekly_workday.Add(true);
                            ACC_workday_from_Sunday++;
                        }
                        remaining_workday_from_Sunday_to.Add(dateime, ACC_workday_from_Sunday);
                        break;
                }
            }
            if (end.DayOfWeek != DayOfWeek.Saturday)
            {
                weekly_workday_list.Add(ACC_workday_from_Sunday);
                yearweek_list.Add(weekofyear_temp);

                // record workday_datetime_till_last_day_of_week
                // updating by looping datetime_thisweek
                int temp_workday = 0;
                int workday_to_Saturday;
                for (int index = 0; index < datetime_thisweek.Count; index++)
                {
                    DateTime dt = datetime_thisweek[index];
                    if (is_weekly_workday[index])
                    {
                        workday_to_Saturday = ACC_workday_from_Sunday - temp_workday;
                        temp_workday++;
                    }
                    else
                    {
                        workday_to_Saturday = ACC_workday_from_Sunday - temp_workday;
                    }
                    remaining_workday_to_Saturday_from.Add(dt, workday_to_Saturday);
                }
            }
        }

        static public int GetStartWeek()
        {
            return GetYearAndWeekOfYear(YearWeek.StartDate);
        }

        static public int GetEndWeek()
        {
            return GetYearAndWeekOfYear(YearWeek.EndDate);
        }

        static public int GetStartWeekIndex()
        {
            return IndexOf(StartDate);
        }

        static public int GetEndWeekIndex()
        {
            return IndexOf(EndDate);
        }

        static public int IndexOf(DateTime datetime)
        {
            int ret_index = invalid_index;
            if ((datetime >= StartDate) && (datetime <= EndDate))
            {
                ret_index = IndexOf(GetYearAndWeekOfYear(datetime));
            }
            return ret_index;
        }

        static public int IndexOf(int year_and_week)
        {
            int ret_index = invalid_index;
            if (yearweek_list.Contains(year_and_week))
            {
                ret_index = yearweek_list.IndexOf(year_and_week);
            }
            return ret_index;
        }

        static public int ElementAt(int index)
        {
            int ret_yearweek = invalid_yearweek;
            if ((index >= 0) && (index < yearweek_list.Count))
            {
                ret_yearweek = yearweek_list[index];
            }
            return ret_yearweek;
        }

        static public Boolean IsYearWeekValueInRange(int yearweek_to_check)
        {
            return yearweek_list.Contains(yearweek_to_check);
        }

        // to be implemented -- need to remove working days outside start/end date
        // or need special check for 1st / last week
        static public int GetWorkingDayOfWeekWithinTaskDurationByIndex(int index)
        {
            int ret_weekly_working_day = 0;
            if ((index >= 0) && (index <= (weekly_workday_list.Count - 1)))
            {
                ret_weekly_working_day = weekly_workday_list[index];
            }
            return ret_weekly_working_day;
        }

        static public int GetWorkingDayOfWeekByIndex(int index)
        {
            int ret_weekly_working_day = 0;
            if ((index >= 0) && (index <= (weekly_workday_list.Count - 1)))
            {
                ret_weekly_working_day = weekly_workday_list[index];
            }
            return ret_weekly_working_day;
        }

        static public int GetWorkingDayOfWeekByYearWeek(int year_week)
        {
            int ret_weekly_working_day = 0;
            if (IsYearWeekValueInRange(year_week))
            {
                ret_weekly_working_day = weekly_workday_list[IndexOf(year_week)];
            }
            return ret_weekly_working_day;
        }

        //
        // Input: Date
        // Output: YWW, example 345
        // 
        // NOTE: because we use the week containing 1/1 as first week, we need to check if 1/1 is within this week
        //
        // If 1/1 is within this week, weekno is always 01, otherwise weekno is GetWeekOfYear(CalendarWeekRule.FirstDay)
        //
        //
        static public int GetYearAndWeekOfYear(DateTime datetime)
        {
            CultureInfo my_culture = DateOnly.datetime_culture;
            Calendar my_calendar = my_culture.Calendar;
            DateTimeFormatInfo my_dt_format = my_culture.DateTimeFormat;
            int weekno = my_calendar.GetWeekOfYear(datetime, CalendarWeekRule.FirstDay, my_dt_format.FirstDayOfWeek);
            int yearno = datetime.Year % 10;
            // Special case
            // Detect if 1/1 is the same week as datetime(12/xx)
            // In this case, yearno is last year but weekno  is 1
            // if before December 25(inclusive), skip the rest of special-check
            // 1, 31, 30, 29, 28, 27, 26 of December
            if (((int)datetime.Month == 12) && ((int)datetime.Day > 26))
            {
                int day = datetime.Day;
                int dow = (int)datetime.DayOfWeek;
                int Saturday_of_this_week = day + (6 - dow);
                if (Saturday_of_this_week >= 32) // already January on Saturday of this week
                {
                    weekno = 1;
                    yearno++;
                }
            }
            return (yearno * 100 + weekno);
        }

    }

    static public class DateOnly
    {
        static public DateTime DateTime_Earliest = new DateTime(1900, 1, 1);
        static public DateTime DateTime_Latest = new DateTime(9999, 12, 31);
        static private String cultureName = "en-US";// { "en-US", "ru-RU", "ja-JP" };
        static public CultureInfo datetime_culture = new CultureInfo(cultureName);

        static private DateTime[] active_holiday;
        static private DateTime[] HolidaysSince2023 = 
        {
            // National Holiday 
            new DateTime(2023,  1,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  1, 21, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 22, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 23, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 24, 0, 0, 0),    // CNY
            new DateTime(2023,  2, 28, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  4,  4, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  4,  5, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  5,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  6, 22, 0, 0, 0),    // dragon-boat
            new DateTime(2023,  9, 29, 0, 0, 0),    // mid-autumn
            new DateTime(2023, 10, 10, 0, 0, 0),    // fixed-date holiday

            new DateTime(2024,  1,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  2, 28, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  4,  4, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  4,  5, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  5,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024, 10, 10, 0, 0, 0),    // fixed-date holiday
            // National Holiday on weekend -- shifted off
            new DateTime(2023,  1,  2, 0, 0, 0),    // NY
            new DateTime(2023,  1, 25, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 26, 0, 0, 0),    // CNY
            // Company shift off
            new DateTime(2023,  1, 20, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 27, 0, 0, 0),    // CNY
            new DateTime(2023,  2, 27, 0, 0, 0),    // 228
            new DateTime(2023,  4,  3, 0, 0, 0),    // 44&45
            new DateTime(2023,  6, 23, 0, 0, 0),    // dragon-boat
            new DateTime(2022, 10,  9, 0, 0, 0),    // 10*2
            // Typhoon off
            new DateTime(2023,  8,  3, 0, 0, 0),

        };

        // To-be-updated:
        // XM calendar
        static private DateTime[] non_Holiday_weekend_XM = 
        {
            new DateTime(2023,  1,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  1, 21, 0, 0, 0),    // CNY
        };

        static private DateTime[] HolidaysSince2023_XM = 
        {
            // National Holiday 
            new DateTime(2023,  1,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  1, 21, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 22, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 23, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 24, 0, 0, 0),    // CNY
            new DateTime(2023,  2, 28, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  4,  4, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  4,  5, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  5,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2023,  6, 22, 0, 0, 0),    // dragon-boat
            new DateTime(2023,  9, 29, 0, 0, 0),    // mid-autumn
            new DateTime(2023, 10, 10, 0, 0, 0),    // fixed-date holiday

            new DateTime(2024,  1,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  2, 28, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  4,  4, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  4,  5, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024,  5,  1, 0, 0, 0),    // fixed-date holiday
            new DateTime(2024, 10, 10, 0, 0, 0),    // fixed-date holiday
            // National Holiday on weekend -- shifted off
            new DateTime(2023,  1,  2, 0, 0, 0),    // NY
            new DateTime(2023,  1, 25, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 26, 0, 0, 0),    // CNY
            // Company shift off
            new DateTime(2023,  1, 20, 0, 0, 0),    // CNY
            new DateTime(2023,  1, 27, 0, 0, 0),    // CNY
            new DateTime(2023,  2, 27, 0, 0, 0),    // 228
            new DateTime(2023,  4,  3, 0, 0, 0),    // 44&45
            new DateTime(2023,  6, 23, 0, 0, 0),    // dragon-boat
            new DateTime(2022, 10,  9, 0, 0, 0),    // 10*2
            // Typhoon off
            new DateTime(2023,  8,  3, 0, 0, 0),

        };

        static public void SortHoliday()
        {
            Array.Sort<DateTime>(HolidaysSince2023);
            active_holiday = HolidaysSince2023;
        }

        static public T[] SubArray<T>(this T[] data, int index, int length)
        {
            T[] result = new T[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }

        static public void Update_Holiday_Range(DateTime start, DateTime end)
        {
            int start_index = 0;
            int length = HolidaysSince2023.Count();

            start_index = Array.BinarySearch<DateTime>(HolidaysSince2023, start);
            start_index = (start_index >= 0) ? start_index : ~start_index;
            int search_end = Array.BinarySearch<DateTime>(HolidaysSince2023, end);
            search_end = (search_end >= 0) ? search_end : ((~search_end) - 1);
            length = search_end - start_index + 1;

            active_holiday = HolidaysSince2023.SubArray(start_index, length);
        }

        static public Boolean IsHoliday(DateTime datetime)
        {
            Boolean ret = false;

            if ((datetime.DayOfWeek == DayOfWeek.Saturday) || (datetime.DayOfWeek == DayOfWeek.Sunday))
            {
                // Weekend as holiday by default in TWN
                ret = true;
            }
            else
            {
                // if it is a weekday, then check if it is a holiday which is not on weekend
                //foreach (DateTime holiday in HolidaysSince2023)
                //{
                //   if (holiday.Date == datetime.Date)
                //    {
                //        // holiday found, stop checking
                //        ret = true;
                //        break;
                //    }
                //}
                if (Array.BinarySearch<DateTime>(active_holiday, datetime) >= 0)
                {
                    ret = true;
                }
            }

            return ret;
        }

        //static public int BusinessDaysUntil(this DateTime firstDay, DateTime lastDay, params DateTime[] bankHolidays)
        static public int BusinessDaysUntil(this DateTime firstDay, DateTime lastDay)
        {
            DateTime[] bankHolidays = active_holiday;
            firstDay = firstDay.Date;
            lastDay = lastDay.Date;
            if (firstDay > lastDay)
            {
                //throw new ArgumentException("Incorrect last day " + lastDay);
                return -1;
            }

            TimeSpan span = lastDay - firstDay;
            int businessDays = span.Days + 1;
            int fullWeekCount = businessDays / 7;
            // find out if there are weekends during the time exceedng the full weeks
            if (businessDays > fullWeekCount * 7)
            {
                // we are here to find out if there is a 1-day or 2-days weekend
                // in the time interval remaining after subtracting the complete weeks
                //int firstDayOfWeek = (int)firstDay.DayOfWeek;
                //int lastDayOfWeek = (int)lastDay.DayOfWeek;
                int firstDayOfWeek = firstDay.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)firstDay.DayOfWeek;
                int lastDayOfWeek = lastDay.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)lastDay.DayOfWeek;
                if (lastDayOfWeek < firstDayOfWeek)
                    lastDayOfWeek += 7;
                if (firstDayOfWeek <= 6)
                {
                    if (lastDayOfWeek >= 7)// Both Saturday and Sunday are in the remaining time interval
                        businessDays -= 2;
                    else if (lastDayOfWeek >= 6)// Only Saturday is in the remaining time interval
                        businessDays -= 1;
                }
                else if (firstDayOfWeek <= 7 && lastDayOfWeek >= 7)// Only Sunday is in the remaining time interval
                    businessDays -= 1;
            }

            // subtract the weekends during the full weeks in the interval
            businessDays -= fullWeekCount + fullWeekCount;

            // subtract the number of bank holidays during the time interval
            foreach (DateTime bankHoliday in bankHolidays)
            {
                DateTime bh = bankHoliday.Date;
                int day_of_week = (int)bh.DayOfWeek;

                if ((day_of_week > (int)DayOfWeek.Sunday) && (day_of_week < (int)DayOfWeek.Saturday))
                {
                    // reduce one working day if holiday within (firstDay,lastDay) is not on weekend
                    if (firstDay <= bh && bh <= lastDay)
                        --businessDays;
                }
            }

            return businessDays;
        }


        // Get the datetime which is earlier
        static public DateTime ReturnEarlierDateTime(DateTime datetime, String datetime_str)
        {
            if (String.IsNullOrWhiteSpace(datetime_str))
            {
                return datetime.Date;
            }

            try
            {
                DateTime dt = Convert.ToDateTime(datetime_str, datetime_culture).Date;
                if (dt < datetime.Date)
                {
                    datetime = dt.Date;
                }
            }
            catch (Exception ex)
            {
            }
            return datetime.Date;
        }

        // Get the datetime which is later
        static public DateTime ReturnLaterDate(DateTime datetime, String datetime_str)
        {
            if (String.IsNullOrWhiteSpace(datetime_str))
            {
                return datetime.Date;
            }

            try
            {
                DateTime dt = Convert.ToDateTime(datetime_str, datetime_culture).Date;
                if (dt > datetime.Date)
                {
                    datetime = dt.Date;
                }
            }
            catch (Exception ex)
            {
            }
            return datetime.Date;
        }

        static public DateTime FindEearliestTargetStartDate(List<ManPower> manpower)
        {
            //Target_start_date
            DateTime earliest_dt = DateTime_Latest.Date;  // default for no earliest date
            foreach (ManPower mp in manpower)
            {
                earliest_dt = ReturnEarlierDateTime(earliest_dt, mp.Target_start_date);
            }
            return earliest_dt.Date;
        }

        static public DateTime FindLatestTargetEndDate(List<ManPower> manpower)
        {
            //Target_end_date
            DateTime latest_dt = DateTime_Earliest.Date;  // default for no latest date
            foreach (ManPower mp in manpower)
            {
                latest_dt = ReturnLaterDate(latest_dt, mp.Target_end_date);
            }
            return latest_dt.Date;
        }

        static private DateTime FindLatestDueDate(List<ManPower> manpower)
        {
            //Target_end_date
            DateTime latest_dt = DateTime_Earliest.Date;  // default for no latest date
            foreach (ManPower mp in manpower)
            {
                latest_dt = ReturnLaterDate(latest_dt, mp.Due_date);
            }
            return latest_dt.Date;
        }

        static public Boolean IsBetween(DateTime date_checked, DateTime from, DateTime to)
        {
            Boolean b_ret = false;

            if ((date_checked.Date >= from.Date) && (date_checked.Date <= to.Date))
            {
                b_ret = true;
            }

            return b_ret;
        }
    }

    public class Site
    {
        // static for class Site
        static private int init_value = -1;
        static private String[] SiteListString = { "Undefined", "HQ", "XM" };
        static private List<String> InternalSiteList = SiteListString.ToList();
        static public List<String> SiteList = InternalSiteList.GetRange(1, 2);
        static public int Count = SiteList.Count;

        // internal variable
        private int site = init_value;

        // member function
        public int Index   // property
        {
            get { return site; }    // get method
            set                     // set method
            {
                if ((value >= 0) && (value < Count))
                {
                    site = value;
                }
                else
                {
                    site = init_value;
                }
            }
        }

        public String Name   // property
        {
            get
            {
                if ((site >= 0) && (site < Count))
                {
                    return SiteListString[site + 1];
                }
                else
                {
                    return SiteListString[0];
                }
            }   // get method
            set { site = SiteList.IndexOf(value.Substring(0, 2)); }  // set method
        }
    }

    public class ManPowerDate
    {
        static public DateTime Earliest = new DateTime(1900, 1, 1);
        static public DateTime Latest = new DateTime(9999, 12, 31);
        static private String CultureName = "en-US";// { "en-US", "ru-RU", "ja-JP" };
        static public CultureInfo CultureInfo = new CultureInfo(CultureName);

        private DateTime date;
        public DateTime Date                        // property
        {
            get { return date.Date; }               // get method
            set { date = value.Date; }            // set method
        }

        public ManPowerDate() { this.date = Earliest; }
        public ManPowerDate(DateTime date) { this.date = date; }

        public Boolean IsHoliday(ManPowerHolidayList holidays)
        {
            int index = holidays.IndexOf(this);
            return (index >= 0);
        }
        public Boolean IsBetween(ManPowerDate from, ManPowerDate to)
        {
            if ((Compare(from, this) >= 0) && (Compare(this, to) <= 0))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public ManPowerDate ReturnEarlier(ManPowerDate date)
        {
            ManPowerDate ret_date = (Compare(this, date) > 0)? date: this;
            return ret_date;
        }
        public ManPowerDate ReturnLater(ManPowerDate date)
        {
            ManPowerDate ret_date = (Compare(this, date) < 0) ? date : this;
            return ret_date;
        }
        public void FromString(String date_string)
        {
            date = DateTime.Parse(date_string);
        }
        public List<ManPowerDate> ToList()
        {
            List<ManPowerDate> ret_list = new List<ManPowerDate>();
            ret_list.Add(this);
            return ret_list;
        }

        static public int Compare(ManPowerDate first_date, ManPowerDate second_date)
        {
            DateTime d1 = first_date.Date, d2 = second_date.Date;
            int compare_result = DateTime.Compare(d1, d2);
            return compare_result;
        }

    }

    public class ManPowerDateComparer : IComparer<ManPowerDate>
    {
        public int Compare(ManPowerDate x, ManPowerDate y)
        {
            return ManPowerDate.Compare(x, y);
        }
    }

    public class ManPowerHolidayList
    {
        public List<ManPowerDate> Holidays = new List<ManPowerDate>();
        public Site Site;
        public void Add(ManPowerDate date)
        {
            Holidays.Add(date);
            Holidays.Sort(ManPowerDate.Compare);
        }
        public void AddRange(List<ManPowerDate> date_list)
        {
            Holidays.AddRange(date_list);
            Holidays.Sort(ManPowerDate.Compare);
        }
        public int IndexOf(ManPowerDate date)
        {
            return Holidays.IndexOf(date);
        }
        // list holidays between
        public int OffDayBetween(ManPowerDate firstDay, ManPowerDate lastDay)
        {
            ManPowerDateComparer date_compare = new ManPowerDateComparer();
            int index_from = Holidays.BinarySearch(firstDay, date_compare);
            int index_to = Holidays.BinarySearch(lastDay, date_compare);

            // update index result
            index_from = (index_from >= 0) ? index_from : ~index_from;
            index_to = (index_to >= 0) ? index_to : (~index_to) - 1;

            int day_count = (index_to - index_from + 1);
            return day_count;
        }
        public int BussinessDayBetween(ManPowerDate firstDay, ManPowerDate lastDay)
        {
            TimeSpan span = lastDay.Date - firstDay.Date;
            int day_count = (span.Days + 1) - OffDayBetween(firstDay,lastDay);

            return day_count;
            //firstHoliday = (index_from >= 0) ? Holidays[index_from] : Holidays[~index_from];
            //lastHoliday = (index_to >= 0) ? Holidays[index_to] : Holidays[(~index_to) - 1];
        }
        // The zero-based index of item in the sorted List<T>, if item is found; 
        // otherwise, a negative number that is the bitwise complement of the index of the next element that is larger than item or, 
        // if there is no larger element, the bitwise complement of Count.
        // Bitwise complement operator ~
    }
}
