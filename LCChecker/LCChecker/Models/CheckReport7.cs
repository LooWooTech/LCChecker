using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckReport7 : CheckEngine,ICheck
    {


        public CheckReport7(string filePath,List<Project> Projects)
        {
            GetMessage(filePath);
            var list = new List<IRowRule>();
            Dictionary<string, Project> Team = new Dictionary<string, Project>();
            foreach (var item in Projects)
            {
                Team.Add(item.ID, item);
            }

            list.Add(new OnlyProject() { ColumnIndex = 3, Values = new[] { "项目编号", "市", "县", "项目名称", "新增耕地面积" } });
            list.Add(new SpecialData() { ColumnIndex = 7, Value = "耕地质量等别", IDIndex = 3, ProjectData = Ship });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" } });

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }


    }
}