using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Models
{
    public class CheckReport5 : CheckEngine,ICheck
    {

        public CheckReport5(List<Project> projects)
        {
            SetWhether(projects);
            var list = new List<IRowRule>();
            foreach (var item in projects)
            {
                var rule1 = new StringEqual() { ColumnIndex = 3, Data = item.ID };
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new AndRule()
                    {
                        Rule1 = new StringEqual() { ColumnIndex = 1, Data = item.City.ToString() },
                        Rule2 = new StringEqual() { ColumnIndex=2,Data=item.County}
                    }
                });
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new StringEqual() { ColumnIndex=4,Data=item.Name}
                });
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" } });

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }


        public override void SetWhether(List<Project> projects)
        {
            foreach (var item in projects)
            {
                Whether.Add(item.ID, item.IsHasError);
            }
        }

      
    }
}