using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Models
{
    public class CheckReport4 : CheckEngine,ICheck
    {
        public CheckReport4(string filePath)
        {
            GetMessage(filePath);
            var list = new List<IRowRule>();
            int count = Ship.Count();
            string[] IDS = new string[count];
            int i = 0;
            foreach (var item in Ship.Keys)
            {
                IDS[i] = item;
                i++;
            }

            list.Add(new CellRangeRowRule() { ColumnIndex = 3, Values = IDS });
            foreach (var item in Ship.Keys)
            {
                var rule1 = new StringEqual() { ColumnIndex = 3, Data = item };
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new AndRule()
                    {
                        Rule1 = new StringEqual() { ColumnIndex = 1, Data = Ship[item].City },
                        Rule2 = new StringEqual() { ColumnIndex=2,Data=Ship[item].County}
                    }
                });
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 4, Values = new[] { "重复备案项目", "增减挂钩项目误备案至农村土地整治检测监管系统", "经核实，项目由于___原因未实施或未终止实施，详细说明具体情况" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是", "否" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是", "否" } });
            
            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

    }
}