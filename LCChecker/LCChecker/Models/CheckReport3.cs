using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckReport3:CheckEngine
    {
        public Dictionary<string, string> Ship = new Dictionary<string, string>();

        public CheckReport3(string filePath)
        {
            var list = new List<IRowRule>();
            list.Add(new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是", "否" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是", "否" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" } });

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }


 


        
    }
}