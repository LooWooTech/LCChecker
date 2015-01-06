﻿using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Areas.Second
{
    public class CheckThree:SecondCheckEngine,ISeCheck,ISePlanCheck
    {
        public CheckThree(List<SecondProject> projects) {
            Whether = projects.ToDictionary(e => e.ID, e => e.IsApplyDelete);
            Dictionary<string, SecondProject> Team = projects.ToDictionary(e => e.ID, e => e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, AreaIndex = 2, NewAreaIndex = 3, Projects = Team, Values = new[] { "市", "县", "项目名称", "项目规模", "新增耕地面积" }, ID = "2301（基本规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2302（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" }, ID = "2303（填写规则）" });

            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

        public CheckThree(List<pProject> projects)
        {
            //Whether = projects.ToDictionary(e => (e.Name.Trim().ToUpper() + '-' + e.County.Trim().ToUpper() + '-' + e.Key.Trim().ToUpper()), e => e.IsApplyDelete);
            foreach (var item in projects) {
                if (item.IsApplyDelete) {
                    var key = item.Name.Trim().ToUpper() + '-' + item.County.Trim().ToUpper() + '-' + item.Key.Trim().ToUpper();
                    if (PlanIDS.ContainsKey(key))
                    {
                        PlanIDS[key]++;
                    }
                    else {
                        PlanIDS.Add(key, 1);
                    }
                }
            }
            var list = new List<IRowRule>();
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2302（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" }, ID = "2303（填写规则）" });

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }
    }
}