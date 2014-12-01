using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Areas.Second
{
    public class CheckTwo:SecondCheckEngine,ISeCheck
    {
        public CheckTwo(List<SecondProject> projects) {
            Whether = projects.ToDictionary(e => e.ID, e => !(e.IsHasDoubt||e.IsApplyDelete || e.IsHasError||e.IsPacket||e.IsDescrease||e.IsRelieve));
            Dictionary<string, SecondProject> Team = projects.ToDictionary(e => e.ID, e => e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex=3,AreaIndex=2,NewAreaIndex=3,Projects=Team,ID="2201",Values=new[]{"市","县","项目名称","项目规模","新增耕地面积"}});
            list.Add(new SumRowRule() { SumColumnIndex = 6, ColumnIndices = new[] { 7, 8 } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" } ,ID="2202"});

            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }
    }
}