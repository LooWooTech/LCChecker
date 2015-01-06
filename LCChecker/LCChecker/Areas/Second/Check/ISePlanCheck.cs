using LCChecker.Areas.Second.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LCChecker.Areas.Second
{
    public interface ISePlanCheck
    {
        Dictionary<string, List<string>> GetError();
       // Dictionary<string, int> GetPlanIDS();

        List<pProject> GetPlanData();
        int GetNumber();
        bool Check(string FilePath, ref string Mistakes, SecondReportType Type, bool IsPlan);
    }
}
