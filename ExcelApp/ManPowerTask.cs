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
        public ManPowerDate Task_Start_Date; // Generate according to Target_start_date & Target_end_date;
        public ManPowerDate Task_End_Date;
        public ManPowerDate Task_Due_Date;
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
        static public String CSVCaptionLine;              // reading from CSV
        static public ManPowerDate ManPowerStartDate, ManPowerEndDate;    // search all CSV
        static public int ManPowerStartWeek, ManPowerEndWeek;         // search all CSV
        //static public Boolean IsWorkingDay(String assignee) { }
        //static public String Title_StartDate_to_EndDate;  // Generated according to Start_Date, End_Date
        static public String Title_StartWeek_to_EndWeek;  // Generated according to Start_Date, End_Date
        static public String Title_Project_Action_Owner_WeekOfYear_ManHour;

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

        private void Process_ManPower_Data()
        {
            Task_Project_Name = ManPower.Recent_Task_Project_Name;
            Task_Action_Name = Title;
            Task_Owner_Name = Assignee;
            Task_Start_Date = new ManPowerDate(DateTime.Now);
            Task_End_Date = Task_Start_Date;
            Task_Start_Week = Task_Start_Date.YearWeekNo();
            Task_End_Week = Task_Start_Week;

            // man-power plan needs to be checked and updated later in this function
            Daily_Average_ManHour_string = empty_average_manhour;
            Daily_Average_Manhour_value = empty_average_manhour_value;
            ManHour = -1;

            if (String.IsNullOrWhiteSpace(Target_start_date) == true)     // empty
                return;

            Task_Start_Date = new ManPowerDate(Target_start_date);
            Task_Start_Week = Task_Start_Date.YearWeekNo();

            if (String.IsNullOrWhiteSpace(Target_end_date) == true)     // empty
                return;

            Task_End_Date = new ManPowerDate(Target_end_date);
            Task_End_Week = Task_End_Date.YearWeekNo();

            if (Task_End_Date < Task_Start_Date)
                return;

            Site assignee_site = Site.OfAssignee(Assignee);
            int siteIndex = assignee_site.Index;
            ManPowerHolidayList holidayInUse = ManPowerTask.AllHolidayList[siteIndex];

            if (Double.TryParse(Man_hour, out ManHour) == false)
                return;

            ManPowerDate start = Task_Start_Date;
            ManPowerDate end = Task_End_Date;

            int workday_count = holidayInUse.BussinessDayBetween(start, end);

            // workday must be > 0 (ie, from start to end date shouldn't be all in the middle of holidays)
            // if ==0 set it to 1 (assuemd she/he works 1-day on holiday
            if (workday_count == 0)
            {
                workday_count = 1;
            }

            Double average_man_hour = Math.Round(ManHour / workday_count, ManPower.average_rounding_digit);
            Daily_Average_ManHour_string = average_man_hour.ToString(ManPower.pSpecifier);
            Daily_Average_Manhour_value = average_man_hour;
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
        static public String GenerateDateTitle(ManPowerDate start, ManPowerDate end)
        {
            String ret_str = "";
            if (start > end)
            {
                // to-check: shouldn't be here
            }
            else
            {
                // At least one date (start_date)
                ManPowerDate dt = start;
                ret_str = dt.ToString("d", ManPowerDate.CultureInfo);
                dt++;
                // add "," + next-date till next-date is the end-date
                while (dt <= end)
                {
                    ret_str += "," + dt.ToString("d", ManPowerDate.CultureInfo);
                    dt++;
                }
                // reaching here when the next-date is after the end-date
            }
            return ret_str;
        }

        static public String GenerateWeekOfYearTitle(ManPowerDate start, ManPowerDate end)
        {
            String ret_str = "";
            if (start > end)
            {
                // to-check: shouldn't be here
            }
            else
            {
                // At least one date (start_date)
                ManPowerDate dt = start;

                //ret_str = dt.ToString("yyyy", ManPowerDate.CultureInfo).Substring(3, 1) + dt.GetYearAndWeekOfYear().ToString();
                ret_str = dt.YearWeekNo().ToString();
                dt += 7;
                // add "," + next-date till next-date is the end-date
                while (dt <= end)
                {
                    //ret_str += "," + dt.ToString("yyyy", ManPowerDate.CultureInfo).Substring(3, 1) + dt.GetYearAndWeekOfYear().ToString();
                    ret_str += "," + dt.YearWeekNo().ToString();
                    dt += 7;
                }
                // reaching here when the next-date is after the end-date
            }
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

        static public List<ManPowerHolidayList> AllHolidayList = new List<ManPowerHolidayList>();
        static public List<YearWeek> AllWeekDateList = new List<YearWeek>();

        static public List<ManPower> ReadManPowerTaskCSV(String csv_filename)
        {
            List<ManPower> ret_manpower_list = new List<ManPower>();
            using (TextFieldParser csvParser = new TextFieldParser(csv_filename))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                // Skip the row with the column names
                ManPower.CSVCaptionLine = csvParser.ReadLine();

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

        static public ManPowerDate FindEearliestTargetStartDate(List<ManPower> manpower)
        {
            //Target_start_date
            ManPowerDate earliest_dt = ManPowerDate.InvalidDate;        // default for no latest date
            foreach (ManPower mp in manpower)
            {
                String date_string = mp.Target_start_date;
                if (String.IsNullOrWhiteSpace(date_string) == false)
                {
                    ManPowerDate checkdate = new ManPowerDate(date_string);
                    earliest_dt = earliest_dt.ReturnEarlier(checkdate);
                }
            }
            return earliest_dt;
        }

        static public ManPowerDate FindLatestTargetEndDate(List<ManPower> manpower)
        {
            //Target_end_date
            ManPowerDate latest_dt = ManPowerDate.InvalidDate;          // default for no latest date        
            foreach (ManPower mp in manpower)
            {
                String date_string = mp.Target_end_date;
                if (String.IsNullOrWhiteSpace(date_string) == false)
                {
                    ManPowerDate checkdate = new ManPowerDate(date_string);
                    latest_dt = latest_dt.ReturnLater(checkdate);
                }
            }
            return latest_dt;
        }

        static public List<ManPower> Processing_DateWeekHoliday(List<ManPower> list_before_post_processing)
        {
            // Generated data for ManPower
            ManPower.ManPowerStartDate = FindEearliestTargetStartDate(list_before_post_processing);
            ManPower.ManPowerEndDate = FindLatestTargetEndDate(list_before_post_processing);
            if (ManPower.ManPowerStartDate > ManPower.ManPowerEndDate)
            {
                LogMessage.CheckFunction("Processing_DateWeekHoliday start/end exception");
            }
            ManPower.Title_StartWeek_to_EndWeek = ManPower.GenerateWeekOfYearTitle(ManPower.ManPowerStartDate, ManPower.ManPowerEndDate);

            // Setup YearWeek variables
            for (int site_index = 0; site_index <= Site.MaxIndex; site_index++)
            {
                YearWeek yearweek = new YearWeek();
                yearweek.SetupByStartDateEndDate(site_index, ManPower.ManPowerStartDate, ManPower.ManPowerEndDate);
                AllWeekDateList.Add(yearweek);
            }

            ManPower.ManPowerStartWeek = ManPower.ManPowerStartDate.YearWeekNo();
            ManPower.ManPowerEndWeek = ManPower.ManPowerEndDate.YearWeekNo();

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
            csv.AppendLine(ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour + ManPower.CSVCaptionLine);

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

                    ManPowerDate first_date = mp.Task_Start_Date;
                    ManPowerDate last_date = mp.Task_End_Date;
                    Double daily_average_manhour_value = mp.Daily_Average_Manhour_value;

                    Site assignee_site = Site.OfAssignee(mp.Assignee);
                    int siteIndex = assignee_site.Index;
                    YearWeek yearWeekInUse = AllWeekDateList[siteIndex];

                    int current_wk_index = yearWeekInUse.IndexOf(mp.Task_Start_Week);
                    ManPowerDate current_date = first_date;
                    ManPowerDate week_end_date = current_date.ThisSaturday();

                    while (current_date <= last_date)
                    {
                        int workingday_this_week = yearWeekInUse.WorkdayToSaturdayFrom(current_date);            // calculation always from current_date

                        // adjust week_end_date to last_date if last_date is on/before this Friday. (i.e/ this week is not complete)
                        if (week_end_date > last_date)
                        {
                            week_end_date = last_date;
                            workingday_this_week -= yearWeekInUse.WorkdayToSaturdayFrom(last_date + 1);
                        }

                        Double weekly_manhour = workingday_this_week * daily_average_manhour_value;
                        int current_yearweek = yearWeekInUse.GetElementAtIndex(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + mp.ToString());

                        current_date = week_end_date + 1;   // Go to next week
                        week_end_date += 7;
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
            csv.AppendLine(ManPower.Title_Project_Action_Owner_WeekOfYear_ManHour + ManPower.CSVCaptionLine);

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
                    String Item_Field_String = ManPower.AddQuoteWithComma(mp.Task_Project_Name);
                    Item_Field_String += ManPower.AddQuoteWithComma(mp.Task_Action_Name);
                    Item_Field_String += ManPower.AddQuoteWithComma(mp.Task_Owner_Name);

                    Site assignee_site = Site.OfAssignee(mp.Assignee);
                    int siteIndex = assignee_site.Index;
                    YearWeek yearWeekInUse = AllWeekDateList[siteIndex];

                    int current_wk_index = yearWeekInUse.IndexOf(mp.Task_Start_Week);
                    Double daily_average_manhour_value = mp.Daily_Average_Manhour_value;
                    ManPowerDate week_start = mp.Task_Start_Date;
                    ManPowerDate week_end = week_start.ThisSaturday();
                    ManPowerDate end = mp.Task_End_Date;

                    while (week_start <= end)
                    {
                        int workday = yearWeekInUse.WorkdayToSaturdayFrom(week_start);      // workday is at the moment (week_start to Saturday)
                        if (end < week_end)                                                 // adjustment workday if end is before this week-ending day (Saturday)
                        {
                            workday -= yearWeekInUse.WorkdayToSaturdayFrom(end + 1);        // workday is at the moment (week_start to end)
                            week_end = end;
                        }

                        Double weekly_manhour = workday * daily_average_manhour_value;
                        int current_yearweek = yearWeekInUse.GetElementAtIndex(current_wk_index);
                        // append year-week
                        String Project_Action_Owner_WeekOfYear_ManHour = Item_Field_String + ManPower.AddQuoteWithComma(current_yearweek.ToString());
                        // append man-hour in this week
                        Project_Action_Owner_WeekOfYear_ManHour += ManPower.AddQuoteWithComma(weekly_manhour.ToString(ManPower.pSpecifier));
                        csv.AppendLine(Project_Action_Owner_WeekOfYear_ManHour + category_string + mp.ToString());

                        week_start = week_end + 1;
                        week_end += 7;
                        current_wk_index++;
                    }
                }
            }

            //after your loop
            File.WriteAllText(Storage.GenerateFilenameWithDateTime(manpower_csv, ".csv"), csv.ToString(), Encoding.UTF8);
        }

        static public List<ManPowerHolidayList> LoadSiteHolidayList()
        {
            String holidayCSV = "CompanyOffDayList.csv";
            List<ManPowerHolidayList> all_holiday_list = ManPowerHolidayList.SetupHolidayListFromCSV(holidayCSV);
            return all_holiday_list;
        }
    }

    public class YearWeek
    {
        static public int invalid_index = -1;
        static public int invalid_yearweek = -1;
        private ManPowerDate startDate = new ManPowerDate();
        private ManPowerDate endDate = new ManPowerDate();
        private List<int> yearWeekNoList = new List<int>();         // list all YearWeek withing start/end range
        private List<int> working_day_list = new List<int>();    // list how many workingday in this week.
        private List<int> workingDay = new List<int>();     // list how many workingday in this week.
        private List<int> workingDayToSaturday = new List<int>();
        private List<int> workingDatFromSunday = new List<int>();
        private Site site;

        public List<int> YearWeekList   // property
        {
            get { return yearWeekNoList; }   // get method
        }
        public ManPowerDate StartDate   // property
        {
            get { return startDate; }   // get method
        }
        public ManPowerDate EndDate   // property
        {
            get { return endDate; }   // get method
        }
        public int StartWeek   // property
        {
            get { return startDate.YearWeekNo(); }   // get method
        }
        public int GetEndWeek   // property
        {
            get { return endDate.YearWeekNo(); }   // get method
        }
        public int StartWeekIndex   // property
        {
            get { return IndexOf(startDate); }   // get method
        }
        public int EndWeekIndex     // property
        {
            get { return IndexOf(endDate); }   // get method
        }
        public int this[int index]
        {
            get { return GetElementAtIndex(index); }          // get method
        }
        public Site Site   // property
        {
            get { return site; }   // get method
        }

        private int IndexOf(ManPowerDate datetime)
        {
            int ret_index = invalid_index;
            if ((datetime >= startDate) && (datetime <= endDate))
            {
                ret_index = yearWeekNoList.IndexOf(datetime.YearWeekNo());
            }
            return ret_index;
        }
        public int IndexOf(int year_and_week)
        {
            int ret_index = yearWeekNoList.IndexOf(year_and_week);
            return ret_index;
        }
        public int GetElementAtIndex(int index)
        {
            int ret_yearweek = invalid_yearweek;
            if ((index >= 0) && (index < yearWeekNoList.Count))
            {
                ret_yearweek = yearWeekNoList[index];
            }
            return ret_yearweek;
        }

        public int WorkdayToSaturdayFrom(ManPowerDate datetime)
        {
            if (datetime.IsBetween(startDate, endDate))
            {
                int diff_day = datetime - startDate;
                return workingDayToSaturday[diff_day];
            }
            else
            {
                return 0;
            }
        }
        public int WorkdayFromSundayTo(ManPowerDate datetime)
        {
            if (datetime.IsBetween(startDate, endDate))
            {
                int diff_day = datetime - startDate;
                return workingDatFromSunday[diff_day];
            }
            else
            {
                return 0;
            }
        }

        // to update -- become site-dependent
        public void SetupByStartDateEndDate(int siteindex, ManPowerDate start, ManPowerDate end)
        {
            startDate = start;
            endDate = end;
            // List to be created
            yearWeekNoList.Clear();
            working_day_list.Clear();
            workingDay.Clear();
            workingDayToSaturday.Clear();
            workingDatFromSunday.Clear();
            // END

            site = new Site(siteindex);
            ManPowerHolidayList holiday_in_use = ManPowerTask.AllHolidayList[siteindex];

            if (end < start)
            {
                // shouldn't be here
                return;
            }

            ManPowerDate current_date = start;
            ManPowerDate week_end = current_date.ThisSaturday();

            while (current_date <= end)
            {
                int yearweekno = current_date.YearWeekNo();
                yearWeekNoList.Add(yearweekno);

                week_end = (week_end > end) ? end : week_end;               // Should be Saturday or "end"
                int workingdayThisWeek;
                workingdayThisWeek = holiday_in_use.BussinessDayBetween(current_date, week_end);
                workingDay.Add(workingdayThisWeek);

                int acc_workingday = 0;
                // iterate for this week (or until "end" date)
                while (current_date <= week_end)
                {
                    workingDayToSaturday.Add(workingdayThisWeek);
                    if (current_date.IsWorkingday(holiday_in_use))
                    {
                        acc_workingday++;
                        workingdayThisWeek--;
                    }
                    workingDatFromSunday.Add(acc_workingday);
                    current_date++;
                }
                // week_start should be Sunday(week_end + 1) or "end" now
                week_end += 7;                                  // Should be next Saturday or "end"+7
            }
        }

        public Boolean IsYearWeekValueInRange(int yearweek_to_check)
        {
            return yearWeekNoList.Contains(yearweek_to_check);
        }

        // to be implemented -- need to remove working days outside start/end date
        // or need special check for 1st / last week
        public int GetWorkingDayOfWeekWithinTaskDurationByIndex(int index)
        {
            int ret_weekly_working_day = 0;
            if ((index >= 0) && (index <= (workingDay.Count - 1)))
            {
                ret_weekly_working_day = workingDay[index];
            }
            return ret_weekly_working_day;
        }

        public int GetWorkingDayOfWeekByIndex(int index)
        {
            int ret_weekly_working_day = 0;
            if ((index >= 0) && (index <= (workingDay.Count - 1)))
            {
                ret_weekly_working_day = workingDay[index];
            }
            return ret_weekly_working_day;
        }

        public int GetWorkingDayOfWeekByYearWeek(int year_week)
        {
            int ret_weekly_working_day = 0;
            if (IsYearWeekValueInRange(year_week))
            {
                ret_weekly_working_day = workingDay[IndexOf(year_week)];
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
    }

    public class Site
    {
        // static for class Site
        static public int UndefinedSiteIndex = -1;
        static public String UndefinedSite = "UndefinedSite";
        static private String[] siteNameList = { "HQ", "XM" };
        static public List<String> SiteList = siteNameList.ToList();
        static private int defaultSiteIndex = 0;        // "HQ" as default

        // internal variable
        static private int init_value = UndefinedSiteIndex;
        private int site = init_value;

        static private String[] site_assignee_pair = 
        { 
            "Kiara Zou 鄒寶鳳", "XM",
            "Xiaohong Chen 陈晓红", "XM",
            "Harry HH Lin 林鸿晖", "XM",
            "Linjue Lu 盧林覺", "XM",
            "WanqiaoXie 谢万桥", "XM",
            "QiuPing Deng 鄧秋萍", "XM",
            "Jianhang Chen 陈建行", "XM",
            "Sally Huang 黄思婷", "XM" 
        };

        // member function
        public Site() { }
        public Site(int site_index) { Index = site_index; }
        public Site(String site_name) { Name = site_name; }
        public List<Site> ToList() { List<Site> site_list = new List<Site>(); site_list.Add(this); return site_list; }
        private Boolean SiteIndexInRange(int index)
        {
            return ((index >= 0) && (index < SiteList.Count)) ? true : false;
        }
        static public int MaxIndex   // property
        {
            get { return SiteList.Count - 1; }    // get method
        }
        public int Index   // property
        {
            get { return site; }    // get method
            set { site = SiteIndexInRange(value) ? value : UndefinedSiteIndex; }   // set method
        }
        public String Name   // property
        {
            get { return SiteIndexInRange(site) ? siteNameList[site] : UndefinedSite; }    // get method
            set { site = ((String.IsNullOrWhiteSpace(value) == true) || (value.Length < 2)) ? UndefinedSiteIndex : SiteList.IndexOf(value.Substring(0, 2)); }   // set method // check if empty or string too short
        }
        static public Site OfAssignee(String assignee)
        {
            List<String> site_assignee_list = site_assignee_pair.ToList<String>();
            int index = site_assignee_list.IndexOf(assignee);
            if (index >= 0)
            {
                String sitename = site_assignee_list[index + 1];
                return new Site(sitename);
            }
            else
            {
                // if assignee is not on the list, use default Site
                return new Site(defaultSiteIndex);
            }
        }
        private int SiteIndexByName(String siteName)
        {
            int index = SiteList.IndexOf(siteName);
            return (index >= 0) ? index : UndefinedSiteIndex;
        }

        public static Boolean operator ==(Site a, Site b)
        {
            return (a.site == b.site) ? true : false;
        }
        public static Boolean operator !=(Site a, Site b)
        {
            return !(a == b);
        }

    }

    public class ManPowerDate
    {
        static private DateTime earliest = new DateTime(1900, 1, 1);
        static private DateTime latest = new DateTime(9999, 12, 31);
        static public ManPowerDate InvalidDate = new ManPowerDate(earliest);
        static public ManPowerDate Earliest = new ManPowerDate(earliest.AddDays(1.0));
        static public ManPowerDate Latest = new ManPowerDate(latest);
        static private String CultureName = "en-US";// { "en-US", "ru-RU", "ja-JP" };
        static public CultureInfo CultureInfo = new CultureInfo(CultureName);
        static public Calendar Calendar = CultureInfo.Calendar;
        static public DateTimeFormatInfo DateTimeFormatInfo = CultureInfo.DateTimeFormat;
        static private ManPowerDateComparer date_compare = new ManPowerDateComparer();

        private DateTime date;
        public DateTime Date                        // property
        {
            get { return date.Date; }               // get method
            set { date = value.Date; }            // set method
        }

        public DayOfWeek DayOfWeek   // property
        {
            get { return date.DayOfWeek; }   // get method
            //set { count[(int)SeverityOrder.B] = value; }  // set method
        }

        public ManPowerDate() { this.date = earliest; }
        public ManPowerDate(DateTime date) { this.date = date; }
        public ManPowerDate(String date_string) { FromString(date_string); }

        public Boolean IsHoliday(ManPowerHolidayList holidays)
        {
            Boolean is_holiday = holidays.IsHoliday(this);
            return is_holiday;
        }
        public Boolean IsWorkingday(ManPowerHolidayList holidays)
        {
            Boolean is_holiday = holidays.IsHoliday(this);
            return !is_holiday;
        }
        public Boolean IsBetween(ManPowerDate from, ManPowerDate to)
        {
            if (from == InvalidDate)
            {
                return false;
            }
            if (to == InvalidDate)
            {
                return false;
            }

            if ((this >= from) && (from <= to))
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
            if (date == InvalidDate)
                return this;

            if (this == InvalidDate)
                return date;

            ManPowerDate ret_date = (this < date) ? this : date;
            return ret_date;
        }
        public ManPowerDate ReturnLater(ManPowerDate date)
        {
            if (date == InvalidDate)
                return this;

            if (this == InvalidDate)
                return date;

            ManPowerDate ret_date = (this > date) ? this : date;
            return ret_date;
        }
        private void FromString(String date_string)
        {
            if (String.IsNullOrWhiteSpace(date_string) == true)     // empty
                date = InvalidDate.date;

            CultureInfo cultureInfo = new CultureInfo("en-GB");
            try
            {
                date = Convert.ToDateTime(date_string, cultureInfo);
            }
            catch (Exception ex)
            {
                date = InvalidDate.date;
            }
        }
        public String ToString(String str) { return date.ToString(str); }
        public String ToString(IFormatProvider format) { return date.ToString(format); }
        public String ToString(String str, IFormatProvider format)
        {
            String ret_str = date.ToString(str, format);
            return ret_str;
        }
        public List<ManPowerDate> ToList()
        {
            List<ManPowerDate> ret_list = new List<ManPowerDate>();
            ret_list.Add(this);
            return ret_list;
        }
        public int YearWeekNo()
        {
            DateTime datetime = this.date;
            int weekno = Calendar.GetWeekOfYear(datetime, CalendarWeekRule.FirstDay, DateTimeFormatInfo.FirstDayOfWeek);
            int yearno = datetime.Year % 10;

            // if last week of year, need to check if 1/1 is the same week.
            // if yes, yearno++ weekno=1
            if (weekno == 53)
            {
                if (this.ThisSaturday().date.Month == 1)
                {
                    yearno++;
                    weekno = 1;
                }
            }
            return (yearno * 100 + weekno);
        }
        public ManPowerDate ThisSaturday()
        {
            if (this != InvalidDate)
            {
                int days = 6 - (int)date.DayOfWeek;
                return (this + days);
            }
            else
            {
                return this;
            }
        }
        public ManPowerDate ThisSunday()
        {
            if (this != InvalidDate)
            {
                int days = (int)date.DayOfWeek;
                return (this - days);
            }
            else
            {
                return this;
            }
        }

        public static Boolean operator <(ManPowerDate a, ManPowerDate b)
        {
            if (date_compare.Compare(a, b) < 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static Boolean operator ==(ManPowerDate a, ManPowerDate b)
        {
            if (date_compare.Compare(a, b) == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static Boolean operator <=(ManPowerDate a, ManPowerDate b)
        {
            if (date_compare.Compare(a, b) <= 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static Boolean operator >(ManPowerDate a, ManPowerDate b)
        {
            return (b < a);
        }
        public static Boolean operator >=(ManPowerDate a, ManPowerDate b)
        {
            return (b <= a);
        }
        public static Boolean operator !=(ManPowerDate a, ManPowerDate b)
        {
            return !(a == b);
        }
        public static ManPowerDate operator +(ManPowerDate a, Double b)
        {
            if (a != InvalidDate)
            {
                return new ManPowerDate(a.date.AddDays(b));
            }
            else
            {
                return a;
            }
        }
        public static ManPowerDate operator +(ManPowerDate a, int b)
        {
            return (a + (Double)b);
        }
        public static ManPowerDate operator -(ManPowerDate a, int b)
        {
            return (a + (-b));
        }
        public static ManPowerDate operator -(ManPowerDate a, Double b)
        {
            return (a + (-b));
        }
        public static int operator -(ManPowerDate a, ManPowerDate b)
        {
            if (a == InvalidDate)
            {
                return 0;
            }
            if (b == InvalidDate)
            {
                return 0;
            }

            TimeSpan span = a.Date - b.Date;
            int day_count = (span.Days);
            return day_count;
        }
        public static ManPowerDate operator ++(ManPowerDate a)
        {
            return (a + 1);
        }
        public static ManPowerDate operator --(ManPowerDate a)
        {
            return (a - 1);
        }
    }

    public class ManPowerDateComparer : IComparer<ManPowerDate>
    {
        public int Compare(ManPowerDate x, ManPowerDate y)
        {
            DateTime d1 = x.Date, d2 = y.Date;
            int compare_result = DateTime.Compare(d1, d2);
            return compare_result;
        }
    }

    public class ManPowerHolidayList
    {
        private List<ManPowerDate> Holidays = new List<ManPowerDate>();
        private ManPowerDateComparer date_compare = new ManPowerDateComparer();
        public ManPowerHolidayList() { }
        public ManPowerHolidayList(Site site) { Site = site; }
        public Site Site;
        public Boolean IsSite(Site site)
        {
            if (this.Site == site)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public void Add(ManPowerDate date)
        {
            // skip if already exists
            if (IndexOf(date) >= 0)
                return;
            Holidays.Add(date);
            Holidays.Sort(date_compare);
        }
        public void AddRange(List<ManPowerDate> date_list)
        {
            // skip if already exists
            foreach (ManPowerDate mp_date in date_list)
            {
                if (IndexOf(mp_date) >= 0)
                    continue;
                Holidays.Add(mp_date);
            }
            Holidays.Sort(date_compare);
        }
        public int IndexOf(ManPowerDate date)
        {
            int index = Holidays.BinarySearch(date, date_compare);
            index = (index >= 0) ? index : -1; //
            return index;
        }
        public Boolean IsHoliday(ManPowerDate date)
        {
            if (IndexOf(date) >= 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        // list holidays between
        public int OffDayBetween(ManPowerDate firstDay, ManPowerDate lastDay)
        {
            int day_count = 0;

            // calculation is assumed that (firstDay <= lastDay)
            if (firstDay > lastDay)
            {
                return OffDayBetween(lastDay, firstDay);
            }

            int index_from = Holidays.BinarySearch(firstDay, date_compare);
            if (index_from >= 0)
            {
                // if firstDay is also an holiday
                day_count++;
                index_from++;
            }
            else
            {
                index_from = ~index_from;   // go to next holiday (after firstDay)
            }

            int index_to = Holidays.BinarySearch(lastDay, date_compare);
            if (index_to >= 0)
            {
                // if lastDay is also an holiday
                day_count++;
                index_to--;
            }
            else
            {
                index_to = (~index_to) - 1;   // go to previous holiday (before lastDay)
            }

            day_count += index_to - index_from + 1; // adjusted by holidays between firstDay & lastDay (excluding firstDay/lastDay)
            return day_count;
        }
        public int BussinessDayBetween(ManPowerDate firstDay, ManPowerDate lastDay)
        {
            int day_count = (Math.Abs(lastDay - firstDay) + 1) - OffDayBetween(firstDay, lastDay);

            return day_count;
            //firstHoliday = (index_from >= 0) ? Holidays[index_from] : Holidays[~index_from];
            //lastHoliday = (index_to >= 0) ? Holidays[index_to] : Holidays[(~index_to) - 1];
        }
        // The zero-based index of item in the sorted List<T>, if item is found; 
        // otherwise, a negative number that is the bitwise complement of the index of the next element that is larger than item or, 
        // if there is no larger element, the bitwise complement of Count.
        // Bitwise complement operator ~

        static public List<ManPowerHolidayList> SetupHolidayListFromCSV(String csv_file)
        {
            List<ManPowerHolidayList> ret_list_of_site_holidays = new List<ManPowerHolidayList>();

            // Create a list of site-based empty ManPowerHolidayList (holiday to be added later according to CSV) 
            for (int site_index = 0; site_index < Site.SiteList.Count; site_index++)
            {
                // setup site info
                Site this_site = new Site(site_index);
                // Add holiday-list & associate site info
                ManPowerHolidayList holidays = new ManPowerHolidayList(this_site);
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
                        String site_string = (item.Length >= 2) ? item.Substring(0, 2) : item;
                        Site this_site = new Site(site_string);
                        title_site.Add(this_site);
                    }
                }

                // Fill Holiday List according to title_site

                while (!csvParser.EndOfData)
                {
                    List<String> elements = new List<String>();
                    // Read current line fields, pointer moves to the next line.
                    elements.AddRange(csvParser.ReadFields());
                    int col_index = -1;
                    foreach (String item in elements)
                    {
                        col_index++;
                        if (String.IsNullOrWhiteSpace(item) == true)     // empty
                            continue;

                        // item count more than title count
                        if (col_index >= title_site.Count)
                            break;

                        // get ManPowerDate
                        ManPowerDate mp_date = new ManPowerDate(item);
                        Site which_site = title_site[col_index];
                        ManPowerHolidayList site_holiday = ret_list_of_site_holidays[which_site.Index];
                        site_holiday.Add(mp_date);
                    }
                }
            }

            return ret_list_of_site_holidays;
        }

        static public ManPowerHolidayList SelectHolidayListBySite(List<ManPowerHolidayList> holidays_lst, Site site)
        {
            foreach (ManPowerHolidayList mpl in holidays_lst)
            {
                if (mpl.Site == site)
                    return mpl;
            }
            return holidays_lst[0];

        }
    }
}
