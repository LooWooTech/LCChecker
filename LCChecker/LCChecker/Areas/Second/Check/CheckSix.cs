using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public class CheckSix:SecondCheckEngine,ISeCheck
    {
        public CheckSix(List<SecondProject> projects) {
            Whether = projects.ToDictionary(e => e.ID, e => true);
            Dictionary<string, SecondProject> Team = projects.ToDictionary(e => e.ID, e => e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, NewAreaIndex = 2, Projects = Team, Values = new[] { "市", "县", "项目名称", "新增耕地面积" }, ID = "2601（基本规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2602（填写规则）" });

            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }
    }
}