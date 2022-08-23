using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReportApplication
{
    static class ExcelTest
    {
        public static bool ExcelTestMainTask(String filename)
        {
            String test_filename = FileFunction.GetFullPath(filename);
            if (!FileFunction.Exists(test_filename))
            {
                return false;
            }
            return true;
        }
    }
}
