using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using System.IO;

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
        Man_hour,
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
        public String Man_hour;

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
        public String Daily_Average_ManHour;
        public Double daily_average_manhour_value;
        public String Daily_ManHour_String;
        public String Project_Action_Owner_WeekOfYear_ManHour;

        // global data
        static public String Caption_Line;              // reading from CSV
        static public DateTime Start_Date, End_Date;    // search all CSV
        static public int Start_Week, End_Week;         // search all CSV
        static public List<Boolean> IsWorkingDay = new List<Boolean>();
        static public String Title_StartDate_to_EndDate;  // Generated according to Start_Date, End_Date
        static public String Title_StartWeek_to_EndWeek;  // Generated according to Start_Date, End_Date
        static public String Title_Project_Action_Owner_WeekOfYear_ManHour;
        static public Dictionary<int, int> WorkingDayInWeek = new Dictionary<int, int>();
        static public List<int> WeekOfYearList = new List<int>();

        static public String hierarchy_string_for_project = "Task";
        static public String hierarchy_string_for_action = "Manpower";
        static public String empty_average_manhour = " ";
        static public Double empty_average_manhour_value = -1.0;

        static public String Recent_Task_Project_Name;

        static public String pSpecifier = "F1";   // floating-point with one digit after decimal
 

        public ManPower() { this.SetMemberByString(new List<String>()); }

        public ManPower(List<String> elements)
        {
            this.SetMemberByString(elements);
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
            Man_hour = members[index++];
            if (Hierarchy == hierarchy_string_for_project)
            {
                ManPower.Recent_Task_Project_Name = Title;
            }
            else if (Hierarchy == hierarchy_string_for_action)   // only man-power to calculate average man-hour
            {
                Process_ManPower_Data();
            }
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
                Task_Start_Week = DateOnly.GetYearAndWeekOfYear(Task_Start_Date);
            }

            if (String.IsNullOrWhiteSpace(Target_end_date))
            {
                average_can_be_calculated = false;
            }
            else
            {
                Task_End_Date = Convert.ToDateTime(Target_end_date, DateOnly.datetime_culture).Date;
                Task_End_Week = DateOnly.GetYearAndWeekOfYear(Task_End_Date);
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

            if (average_can_be_calculated)
            {
                DateTime start = Task_Start_Date;
                DateTime end = Task_End_Date;
                int workday_count = DateOnly.BusinessDaysUntil(start, end);
                if (workday_count > 0)
                {
                    Double average_man_hour = Math.Round(ManHour / workday_count, 1);
                    Daily_Average_ManHour = average_man_hour.ToString(ManPower.pSpecifier);
                    daily_average_manhour_value = average_man_hour;
                }
                else
                {
                    // to check:
                    Daily_Average_ManHour = empty_average_manhour;
                    daily_average_manhour_value = empty_average_manhour_value;
                }
            }
            else
            {
                // man-power plan needs to be checked and updated.
                Daily_Average_ManHour = empty_average_manhour;
                daily_average_manhour_value = empty_average_manhour_value;
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

                ret_str = dt.ToString("yyyy", DateOnly.datetime_culture).Substring(3, 1) + DateOnly.GetYearAndWeekOfYear(dt).ToString();
                dt = dt.AddDays(7.0);
                // add "," + next-date till next-date is the end-date
                while (dt <= end.Date)
                {
                    ret_str += "," + dt.ToString("yyyy", DateOnly.datetime_culture).Substring(3, 1) + DateOnly.GetYearAndWeekOfYear(dt).ToString();
                    dt = dt.AddDays(7.0);
                }
                // reaching here when the next-date is after the end-date
            }
            return ret_str;
        }

        // this function is working properly when title start/end date are set up correctly.
        public String GenerateManPowerDailyEffortString()
        {
            String ret_str = "";

            // check if (1) a man-power item (2) Average ManHour is not empty (3) start/end date is not correct
            if ((this.Hierarchy != hierarchy_string_for_action) || (this.Daily_Average_ManHour == empty_average_manhour) ||
                (this.Task_Start_Date > this.Task_End_Date) || (ManPower.Start_Date > ManPower.End_Date))
            {
                // to-check: shouldn't be here
                return ret_str;
            }

            // Find overlay with Task_Start/Task_End -- by default
            DateTime overlay_start = ManPower.Start_Date, overlay_end = ManPower.End_Date;

            // check 1: Task start is later than Man Power Start or not? later one will be the new overlay start_date
            if (this.Task_Start_Date > overlay_start)
            {
                overlay_start = this.Task_Start_Date;
            }

            // check 2: Task end date is earlier than Man Power End date or not? earlier one will be the new overlay end_date
            if (this.Task_End_Date < overlay_end)
            {
                overlay_end = this.Task_End_Date;
            }

            int overlay_start_index = (int)(overlay_start - ManPower.Start_Date).TotalDays,
                overlay_end_index = (int)(overlay_end - ManPower.Start_Date).TotalDays,
                total_end_index = (int)(ManPower.End_Date - ManPower.Start_Date).TotalDays;

            // 1st day is already overlay-date or not? if yes, average-manhour for working day or "0" for holiday
            // if 1st day is not-yet an overlay-date, fill "0"
            if (overlay_start_index == 0)
            {
                ret_str += (ManPower.IsWorkingDay[0]) ? this.Daily_Average_ManHour : "0";
            }
            else
            {
                ret_str += "0";
            }

            int date_index = 1;

            // before overlay
            while (date_index < overlay_start_index)
            {
                ret_str += ", 0";
                date_index++;
            }

            // during overlay -- output average man-hour
            while (date_index <= overlay_end_index)
            {
                ret_str += ", ";
                ret_str += (ManPower.IsWorkingDay[date_index]) ? this.Daily_Average_ManHour : "0";
                date_index++;
            }

            // after overlay
            while (date_index <= total_end_index)
            {
                ret_str += ", 0";
                date_index++;
            }


            //DateTime filling_Date = ManPower.Start_Date;

            //// 1st date is already overlay-date?
            //if (filling_Date == overlay_start)
            //{
            //    ret_str += this.Average_ManHour;    // if yes, output average manhour
            //}
            //else
            //{
            //    ret_str += "0";
            //}
            //filling_Date = filling_Date.AddDays(1.0);

            //// before overlay
            //while(filling_Date < overlay_start) 
            //{
            //    ret_str += ", 0";
            //    filling_Date = filling_Date.AddDays(1.0);
            //}

            //// during overlay -- output average man-hour
            //while(filling_Date <= overlay_end)
            //{
            //    ret_str += ", ";
            //    if (DateOnly.IsHoliday(filling_Date))
            //    {
            //        ret_str += "0";
            //    }
            //    else
            //    {
            //        ret_str += this.Average_ManHour; 
            //    }
            //    filling_Date = filling_Date.AddDays(1.0);
            //}

            //// after overlay
            //while (filling_Date <= ManPower.End_Date)
            //{
            //    ret_str += ", 0";
            //    filling_Date = filling_Date.AddDays(1.0);
            //}

            this.Daily_ManHour_String = ret_str;
            return ret_str;
        }

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
            return_string += this.Man_hour;  // no need to output comma

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
                    ret_manpower_list.Add(manpower);
                }
            }
            return ret_manpower_list;
        }

        static public List<ManPower> Post_Processing(List<ManPower> list_before_post_processing)
        {
            List<ManPower> ret_manpower_list = list_before_post_processing;

            // Generated data
            ManPower.Start_Date = DateOnly.FindEearliestTargetStartDate(ret_manpower_list);
            ManPower.Start_Week = DateOnly.GetYearAndWeekOfYear(ManPower.Start_Date);

            ManPower.End_Date = DateOnly.FindLatestTargetEndDate(ret_manpower_list);
            ManPower.End_Week = DateOnly.GetYearAndWeekOfYear(ManPower.End_Date);

            DateOnly.Update_Holiday_Range(ManPower.Start_Date, ManPower.End_Date);
            ManPower.Title_StartDate_to_EndDate = ManPower.GenerateDateTitle(ManPower.Start_Date, ManPower.End_Date);
            ManPower.Title_StartWeek_to_EndWeek = ManPower.GenerateWeekOfYearTitle(ManPower.Start_Date, ManPower.End_Date);

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

            // Generated data for each task
            foreach (ManPower mp in ret_manpower_list)
            {
                mp.GenerateManPowerDailyEffortString();
            }

            return ret_manpower_list;
        }

        static public List<ManPower> Post_Processing_V2(List<ManPower> original_manpower_list)
        {
            // Update start/end-date and reduce holiday range to save searching time
            ManPower.Start_Date = DateOnly.FindEearliestTargetStartDate(original_manpower_list);
            ManPower.End_Date = DateOnly.FindLatestTargetEndDate(original_manpower_list);
            DateOnly.Update_Holiday_Range(ManPower.Start_Date, ManPower.End_Date);

            // Create a list of holiday_yes_or_no
            // also calculate working day in a week and record it
            ManPower.IsWorkingDay.Clear();
            ManPower.WorkingDayInWeek.Clear();
            int current_weekofyear = DateOnly.GetYearAndWeekOfYear(ManPower.Start_Date.Date);
            int working_day_of_this_week =0;
            for (DateTime dt = ManPower.Start_Date.Date; dt <= ManPower.End_Date.Date; dt = dt.AddDays(1.0))
            {
                switch (dt.DayOfWeek)
                {
                    case DayOfWeek.Sunday:
                        ManPower.IsWorkingDay.Add(false);       // a holiday --> not a working day
                        current_weekofyear = DateOnly.GetYearAndWeekOfYear(dt);
                        working_day_of_this_week = 0;
                        break;
                    case DayOfWeek.Saturday:
                        ManPower.IsWorkingDay.Add(false);       // a holiday --> not a working day
                        ManPower.WorkingDayInWeek.Add(current_weekofyear,working_day_of_this_week);
                        ManPower.WeekOfYearList.Add(current_weekofyear);
                        working_day_of_this_week = 0;
                        break;
                    default:
                        working_day_of_this_week++;
                        ManPower.IsWorkingDay.Add(true);
                        break;
                 }
            }

            if(ManPower.End_Date.Date.DayOfWeek!=DayOfWeek.Saturday)
            {
                // Add the remaining days into week-list
                ManPower.WorkingDayInWeek.Add(current_weekofyear,working_day_of_this_week);
                ManPower.WeekOfYearList.Add(current_weekofyear);
            }

            // Setup Title line
            ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour = "ProjectStage, TestAction, Owner, Week, ManHour,";
            String Empty_Field_String = ",,,,,";

            List<ManPower> ret_manpower_list = new List<ManPower>();

            // add items in week of year
            foreach (ManPower mp in original_manpower_list)
            {
                if (mp.Hierarchy == ManPower.hierarchy_string_for_project)
                {
                    mp.Project_Action_Owner_WeekOfYear_ManHour = Empty_Field_String;
                }
                else if (mp.Hierarchy == ManPower.hierarchy_string_for_action)   // only man-power to calculate average man-hour
                {
                    String Item_Field_String = mp.Task_Project_Name + "," + mp.Task_Action_Name + "," + mp.Task_Owner_Name + ",";
                    int begin_wk_index = ManPower.WeekOfYearList.IndexOf(mp.Task_Start_Week);
                    int end_wk_index = ManPower.WeekOfYearList.IndexOf(mp.Task_End_Week);
                    int current_wk = begin_wk_index;
                    while (current_wk <= end_wk_index)
                    {
                        mp.Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + current_wk.ToString() + ",";
                        Double weekly_manhour = ManPower.WorkingDayInWeek[current_wk] * mp.daily_average_manhour_value;
                        mp.Project_Action_Owner_WeekOfYear_ManHour += weekly_manhour.ToString(ManPower.pSpecifier) + ",";
                        current_wk++;
                    }
                }
                ret_manpower_list.Add(mp);
            }

            return ret_manpower_list;
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

        static public void ProcessManPowerPlan(String manpower_csv)
        {
            List<ManPower> manpower_list_before = ReadManPowerTaskCSV(manpower_csv);
            List<ManPower> manpower_list = Post_Processing(manpower_list_before);
            //DateTime manpower_due_date = FindLatestDueDate(manpower_list);

            //before your loop
            var csv = new StringBuilder();

            //csv.AppendLine(ManPower.Caption_Line);
            csv.AppendLine(ManPower.Caption_Line + "," + ManPower.Title_StartDate_to_EndDate);

            //in your loop
            foreach (ManPower mp in manpower_list)
            {
                //var newLine = mp.ToString();
                var newLine = mp.ToString() + "," + mp.Daily_ManHour_String;
                csv.AppendLine(newLine);
            }

            //after your loop
            File.WriteAllText(Storage.GenerateFilenameWithDateTime(manpower_csv, ".csv"), csv.ToString(), Encoding.UTF8);
        }

        static public void ProcessManPowerPlan_V2(String manpower_csv)
        {
            List<ManPower> manpower_list_before = ReadManPowerTaskCSV(manpower_csv);
            List<ManPower> manpower_list = Post_Processing_V2(manpower_list_before);
            //DateTime manpower_due_date = FindLatestDueDate(manpower_list);

            //before your loop
            var csv = new StringBuilder();

            //csv.AppendLine(ManPower.Caption_Line);
            csv.AppendLine(ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour + "," + ManPower.Caption_Line);

            //in your loop
            foreach (ManPower mp in manpower_list)
            {
                //var newLine = mp.ToString();
                var newLine = mp.Project_Action_Owner_WeekOfYear_ManHour + "," + mp.ToString();
                csv.AppendLine(newLine);
            }

            //after your loop
            File.WriteAllText(Storage.GenerateFilenameWithDateTime(manpower_csv, ".csv"), csv.ToString(), Encoding.UTF8);
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
                // Weekend as holiday by default
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

        static public int GetYearAndWeekOfYear(DateTime datetime)
        {
            Calendar my_calendar = datetime_culture.Calendar;
            int ret = my_calendar.GetWeekOfYear(datetime, datetime_culture.DateTimeFormat.CalendarWeekRule, datetime_culture.DateTimeFormat.FirstDayOfWeek);
            return ret;
        }

        //static public int BusinessDaysUntil(this DateTime firstDay, DateTime lastDay, params DateTime[] bankHolidays)
        static public int BusinessDaysUntil(this DateTime firstDay, DateTime lastDay)
        {
            DateTime[] bankHolidays = active_holiday;
            firstDay = firstDay.Date;
            lastDay = lastDay.Date;
            if (firstDay > lastDay)
                throw new ArgumentException("Incorrect last day " + lastDay);

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

}
