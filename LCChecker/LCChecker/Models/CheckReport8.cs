using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Models
{
    public class CheckReport8:CheckEngine,ICheck
    {
        public CheckReport8(string filePath,List<Project> projects)
        {
            GetMessage(filePath);
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
                double NewArea = 0.0;
                if (item.NewArea.HasValue)
                {
                    NewArea = item.NewArea.Value;
                }

                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new DoubleEqual() {ColumnIndex=5,data= NewArea}
                });

            }

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        
        }    
    }
}