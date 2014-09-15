using LCChecker.Rules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckReport5:CheckEngine
    {
        public CheckReport5()
        {
            var list = new List<IRowRule>();
            list.Add(new CellRangeRowRule() { ColumnIndex=8,Values=new[]{"是","否"}});

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }
    }
}