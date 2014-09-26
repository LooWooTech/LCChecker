using LCChecker.Rules;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Models
{
    public class CheckReport8:CheckEngine,ICheck
    {
        public CheckReport8(string filePath,List<Project> projects)
        {
            SetWhether(projects);
            GetMessage(filePath);
            Dictionary<string,Project> Team=new Dictionary<string,Project>();
            foreach(var item in projects)
            {
                Team.Add(item.ID,item);
            }
            var list = new List<IRowRule>();
            list.Add(new OnlyProject() { ColumnIndex = 3, Projects = Team, Values = new[] { "项目编号", "市", "县", "项目名称", "新增耕地面积" },ID="2801" });
            list.Add(new SpecialData() { ColumnIndex = 7,ColumnIndex2=6, Value = "由于建设项目未备案应核减占补平衡指标与已与建设项目预挂钩应核销占补平衡指标", IDIndex = 3 ,ProjectData=Ship,ID="2803"});
            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        
        }

        public override void SetWhether(List<Project> projects)
        {
            foreach (var item in projects)
            {
                Whether.Add(item.ID, item.IsDecrease);
            }
        }
    }
}