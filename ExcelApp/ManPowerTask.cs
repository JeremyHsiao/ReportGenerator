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

        static public String Caption_Line;

        public ManPower() { Hierarchy = ""; }
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

        public String AddQuoteWithComma(String item)
        {
            String return_string = AddComma(AddQuote(item));
            return return_string;
        }

        public String ToString()

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
            return_string += this.Man_hour;

            return return_string;
        }
    }

    static public class ManPowerTask
    {
        static public DateTime DateTime_Earliest = new DateTime(1900, 1, 1);
        static public DateTime DateTime_Latest = new DateTime(9999, 12, 31);
        static public String cultureName = "en-US";// { "en-US", "ru-RU", "ja-JP" };
        static public CultureInfo datetime_culture = new CultureInfo(cultureName);

        static public DateTime[] HolidaysSince2023 = 
        {
            // National Holiday 
            new DateTime(2023,  1,  1, 0, 0, 0),
            new DateTime(2023,  1, 21, 0, 0, 0),
            new DateTime(2023,  1, 22, 0, 0, 0),
            new DateTime(2023,  1, 23, 0, 0, 0),
            new DateTime(2023,  1, 24, 0, 0, 0),
            new DateTime(2023,  2, 28, 0, 0, 0),
            new DateTime(2023,  4,  4, 0, 0, 0),
            new DateTime(2023,  4,  5, 0, 0, 0),
            new DateTime(2023,  5,  1, 0, 0, 0),
            new DateTime(2023,  6, 22, 0, 0, 0),
            new DateTime(2023,  9, 29, 0, 0, 0),
            new DateTime(2023, 10, 10, 0, 0, 0),
            // National Holiday on weekend -- shifted off
            new DateTime(2023,  1,  2, 0, 0, 0),
            new DateTime(2023,  1, 25, 0, 0, 0),
            new DateTime(2023,  1, 26, 0, 0, 0),
            // Typhoon off
            new DateTime(2023,  8,  3, 0, 0, 0),
        };

        static public Boolean IsHoliday(DateTime datetime)
        {
            Boolean ret = false;

            foreach (DateTime holiday in HolidaysSince2023)
            {
                if (holiday == datetime)
                {
                    ret = true;
                    break;
                }
                else if (holiday > datetime)
                {
                    break;
                }
            }
            return ret;
        }

        static public int BusinessDaysUntil(this DateTime firstDay, DateTime lastDay, params DateTime[] bankHolidays)
        {
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
                if (firstDay <= bh && bh <= lastDay)
                    --businessDays;
            }

            return businessDays;
        }

        //static public void ReadManPowerTaskCSV(String csv_filename)
        //{
        //    Excel.Workbook wb;
        //    String new_filename = Storage.GenerateFilenameWithDateTime(Filename: csv_filename, FileExt: ".xlsx");
        //    wb = ExcelAction.OpenCSV(csv_filename);
        //    ExcelAction.CloseCSV_SaveAsExcel(workbook: wb, SaveChanges: true, AsFilename: new_filename);
        //}

        static public List<ManPower> ReadManPowerTaskCSV(String csv_filename)
        {
            List<ManPower> ret_manpower_list = new List<ManPower>();
            ManPower manpower;
            using (TextFieldParser csvParser = new TextFieldParser(csv_filename))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                // Skip the row with the column names
                ManPower.Caption_Line = csvParser.ReadLine();

                while (!csvParser.EndOfData)
                {
                    manpower = new ManPower();
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvParser.ReadFields();
                    int index_count = fields.Count();
                    int index = 0;
                    if (index_count <= 0)
                        continue;
                    if (index < index_count)
                        manpower.Hierarchy = fields[index++];
                    if (index < index_count)
                        manpower.Title = fields[index++];
                    if (index < index_count)
                        manpower.Project = fields[index++];
                    if (index < index_count)
                        manpower.Releases = fields[index++];
                    if (index < index_count)
                        manpower.Team = fields[index++];
                    if (index < index_count)
                        manpower.Assignee = fields[index++];
                    if (index < index_count)
                        manpower.Sprint = fields[index++];
                    if (index < index_count)
                        manpower.Target_start_date = fields[index++];
                    if (index < index_count)
                        manpower.Target_end_date = fields[index++];
                    if (index < index_count)
                        manpower.Due_date = fields[index++];
                    if (index < index_count)
                        manpower.Estimates = fields[index++];
                    if (index < index_count)
                        manpower.Parent = fields[index++];
                    if (index < index_count)
                        manpower.Priority = fields[index++];
                    if (index < index_count)
                        manpower.Labels = fields[index++];
                    if (index < index_count)
                        manpower.Components = fields[index++];
                    if (index < index_count)
                        manpower.Issue_key = fields[index++];
                    if (index < index_count)
                        manpower.Issue_status = fields[index++];
                    if (index < index_count)
                        manpower.Progress = fields[index++];
                    if (index < index_count)
                        manpower.Progress_completed = fields[index++];
                    if (index < index_count)
                        manpower.Progress_remaining = fields[index++];
                    if (index < index_count)
                        manpower.Progress_issue_count_IC = fields[index++];
                    if (index < index_count)
                        manpower.To_do_IC = fields[index++];
                    if (index < index_count)
                        manpower.In_progress_IC = fields[index++];
                    if (index < index_count)
                        manpower.Done_IC = fields[index++];
                    if (index < index_count)
                        manpower.Total_IC = fields[index++];
                    if (index < index_count)
                        manpower.Man_hour = fields[index++];

                    ret_manpower_list.Add(manpower);
                }
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

        // Get the datetime which is earlier
        static private DateTime ReturnEarlierDateTime(DateTime datetime, String datetime_str)
        {
            if (String.IsNullOrWhiteSpace(datetime_str))
            {
                return datetime;
            }

            try
            {
                DateTime dt = Convert.ToDateTime(datetime_str, datetime_culture);
                if (dt < datetime)
                {
                    datetime = dt;
                }
            }
            catch (Exception ex)
            {
            }
            return datetime;
        }

        // Get the datetime which is later
        static private DateTime ReturnLaterDateTime(DateTime datetime, String datetime_str)
        {
            if (String.IsNullOrWhiteSpace(datetime_str))
            {
                return datetime;
            }

            try
            {
                DateTime dt = Convert.ToDateTime(datetime_str, datetime_culture);
                if (dt > datetime)
                {
                    datetime = dt;
                }
            }
            catch (Exception ex)
            {
            }
            return datetime;
        }

        static private DateTime FindEearliestTargetStartDate(List<ManPower> manpower)
        {
            //Target_start_date
            DateTime earliest_dt = DateTime_Latest;  // default for no earliest date
            foreach (ManPower mp in manpower)
            {
                earliest_dt = ReturnEarlierDateTime(earliest_dt, mp.Target_start_date);
            }
            return earliest_dt;
        }

        static private DateTime FindLatestTargetEndDate(List<ManPower> manpower)
        {
            //Target_end_date
            DateTime latest_dt = DateTime_Earliest;  // default for no latest date
            foreach (ManPower mp in manpower)
            {
                latest_dt = ReturnLaterDateTime(latest_dt, mp.Target_end_date);
            }
            return latest_dt;
        }

        static private DateTime FindLatestDueDate(List<ManPower> manpower)
        {
            //Target_end_date
            DateTime latest_dt = DateTime_Earliest;  // default for no latest date
            foreach (ManPower mp in manpower)
            {
                latest_dt = ReturnLaterDateTime(latest_dt, mp.Due_date);
            }
            return latest_dt;
        }

        static public void ProcessManPowerPlan(String manpower_csv)
        {
            List<ManPower> manpower_list = ReadManPowerTaskCSV(manpower_csv);
            DateTime manpower_earliest_target_start_date = FindEearliestTargetStartDate(manpower_list);
            DateTime manpower_latest_target_end_date = FindLatestTargetEndDate(manpower_list);
            //DateTime manpower_due_date = FindLatestDueDate(manpower_list);

            //before your loop
            var csv = new StringBuilder();

            csv.AppendLine(ManPower.Caption_Line);

            //in your loop
            foreach (ManPower mp in manpower_list)
            {
                var newLine = mp.ToString();
                csv.AppendLine(newLine);
            }
          
            //after your loop
            File.WriteAllText(Storage.GenerateFilenameWithDateTime(manpower_csv, ".csv"), csv.ToString(), Encoding.UTF8);
        }
    }
}
