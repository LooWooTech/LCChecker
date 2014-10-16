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
            SetWhether(projects);
            Dictionary<string, Project> Team = new Dictionary<string, Project>();
            foreach (var item in projects)
            {
                Team.Add(item.ID, item);
            }
            var list = new List<IRowRule>();
            list.Add(new OnlyProject() { ColumnIndex = 3, Projects=Team,Values = new[] { "项目编号", "市", "县"},ID="2401" });
            //list.Add(new CellRangeRowRule() { ColumnIndex = 4, Values = new[] { "重复备案项目", "增减挂钩项目误备案至农村土地整治监测监管系统", "经核实，项目由于___原因未实施或未终止实施，详细说明具体情况" },ID="2402" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是", "否" } ,ID="2403"});
            list.Add(new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是", "否" } ,ID="2404"});
            
            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

        public override void SetWhether(List<Project> projects)
        {
            foreach (var item in projects)
            {
                Whether.Add(item.ID, item.IsApplyDelete);
            }


        }
    }
}