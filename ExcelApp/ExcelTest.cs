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
            if (!FileFunction.FileExists(test_filename))
            {
                return false;
            }
            return true;
        }

        public static void ReadSheet_TestPlan()
        {
        }

        public static void GenerateTestReportStructure()
        {
            // open standard test report
            // read Test Plan sheet
            // filtered by Do or Not
            // chekck folder and create it if not exist
            // check excel and copy it if not exist
            // option: remove "Not" folder/file
        }
    }
}
