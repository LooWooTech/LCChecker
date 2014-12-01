using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Areas.Second
{
    public class CheckSeven:SecondCheckEngine,ISeCheck
    {
        public CheckSeven(List<SecondProject> projects) {
            Whether = projects.ToDictionary(e => e.ID, e => e.IsDescrease);
            Dictionary<string, SecondProject> Team = projects.ToDictionary(e => e.ID, e => e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, NewAreaIndex = 2, Projects = Team, ID = "2701", Values = new[] { "市", "县", "项目名称", "新增耕地面积" } });
            list.Add(new SumRowRule() { SumColumnIndex = 6, ColumnIndices = new[] { 7, 8 } ,ID="2702"});

            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }
    }
}