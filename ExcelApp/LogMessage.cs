using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReportApplication
{
    class LogMessage
    {
        static public void CheckFunctionAt(String function, int row)
        {
            String message_line;
            message_line = "Warning: please check " + function + " at " + row.ToString();
            Console.WriteLine(message_line);
            MainForm.SystemLogAddLine(message_line);
        }

        static public void CheckFunction(String function)
        {
            String message_line;
            message_line = "Warning: please check " + function;
            Console.WriteLine(message_line);
            MainForm.SystemLogAddLine(message_line);
        }

        static public void CheckFunctionAtRowColumn(String function, int row, int col)
        {
            String message_line;
            message_line = "Warning: please check " + function + " at (" + row.ToString() + "," + col.ToString() + ")";
            Console.WriteLine(message_line);
            MainForm.SystemLogAddLine(message_line);
        }

        static public void WriteLine(String message)
        {
            Console.WriteLine(message);
            MainForm.SystemLogAddLine(message);
        }
    }
}
