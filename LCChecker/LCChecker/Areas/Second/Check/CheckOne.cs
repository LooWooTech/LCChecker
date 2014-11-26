using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Areas.Second
{
    public class CheckOne:SecondCheckEngine
    {
        public CheckOne(List<SecondProject> projects) {
            Dictionary<string,SecondProject> Team=projects.ToDictionary(e=>e.ID,e=>e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, AreaIndex = 0, NewAreaIndex = 0, Projects = Team,ID="2101", Values = new[] { "项目名称", "市", "县" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是", "否" }, ID = "2102" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是", "否" }, ID = "2103" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" }, ID = "2104" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2105" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" }, ID = "2106" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 10, Values = new[] { "是", "否" }, ID = "2107" });
            list.Add(new ConditionalRowRule()
            {
                Condition = new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是" } },
                Rule = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "否" } }
            });
            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

    }
}