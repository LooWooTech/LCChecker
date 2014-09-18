using LCChecker.Rules;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Models
{
    public class CheckReport4 : CheckEngine,ICheck
    {      
        public CheckReport4(List<Project> projects)
        {
            var list = new List<IRowRule>();
            //int count = projects.Count();
            //string[] IDS = new string[count];
            //int i = 0;
            //foreach (var item in projects)
            //{
            //    IDS[i] = item.ID;
            //    i++;
            //}

            //list.Add(new CellRangeRowRule() { ColumnIndex = 3, Values = IDS });
            foreach (var item in projects)
            {
                var rule = new StringEqual() { ColumnIndex = 3, Data = item.ID };
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule,
                    Rule = new AndRule()
                    {
                        Rule1 = new StringEqual() { ColumnIndex = 1, Data =item.City.ToString() },
                        Rule2 = new StringEqual() { ColumnIndex=2,Data=item.County}
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

        public void SetWhether(List<Project> projects)
        {
            foreach (var item in projects)
            {
                //Whether.Add(item.ID, item.Result);
            }
        }
    }
}