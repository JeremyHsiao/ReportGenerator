using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelReportApplication
{
    static class ExcelTest
    {
        public static bool ExcelTestMainTask(String filename)
        {
            String test_filename = FileFunction.GetFullPath(filename);
            if (!FileFunction.FileExists(test_filename))
            {
                return false;
            }
            TestReport.GenerateTestReportStructure(test_filename);
            return true;
        }
    }
}
