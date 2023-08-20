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
            Console.WriteLine("Warning: please check " + function + " at " + row.ToString());
        }

        static public void CheckFunction(String function)
        {
            Console.WriteLine("Warning: please check " + function);
        }

        static public void CheckFunctionAtRowColumn(String function, int row, int col)
        {
            Console.WriteLine("Warning: please check " + function + " at (" + row.ToString() + "," + col.ToString() + ")");
        }

        static public void WriteLine(String message)
        {
            Console.WriteLine(message);
        }
    }
}
