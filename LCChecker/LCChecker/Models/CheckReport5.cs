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
            Dictionary<string, Project> Team = new Dictionary<string, Project>();
            foreach (var item in projects)
            {
                Team.Add(item.ID, item);
            }
            var list = new List<IRowRule>();
            list.Add(new OnlyProject() { ColumnIndex = 3, Projects = Team, Values = new[] { "项目编号", "市", "县", "项目名称" } });
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