using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Areas.Second
{
    public class CheckFour:SecondCheckEngine,ISeCheck
    {
        public CheckFour(List<SecondProject> projects) {
            Whether = projects.ToDictionary(e => e.ID, e => e.IsHasError);
            Dictionary<string, SecondProject> Team = projects.ToDictionary(e => e.ID, e => e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, AreaIndex = 0, NewAreaIndex = 0, Projects = Team, Values = new[] { "市", "县", "项目名称" },ID="2401（基本规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2402（填写规则）" });

            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }
    }
}