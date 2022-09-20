using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelReportApplication
{
    static class UnderDevelopment
    {
    }

    public static class KeywordReport
    {
        static public List<TestPlanKeyword> ListAllKeyword(List<TestPlan> DoPlan)
        {
            List<TestPlanKeyword> ret = new List<TestPlanKeyword>();
            foreach (TestPlan plan in DoPlan)
            {
                plan.OpenDetailExcel();
                List<TestPlanKeyword> plan_keyword = plan.ListKeyword();
                plan.CloseIssueListExcel();
                if (plan_keyword!=null)
                {
                    ret.AddRange(plan_keyword);
                }
            }
            return ret;
        }
    }

}
