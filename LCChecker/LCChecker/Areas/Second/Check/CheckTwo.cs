using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public class CheckTwo:SecondCheckEngine
    {
        public CheckTwo(List<SecondProject> projects) {
            Dictionary<string, SecondProject> Team = projects.ToDictionary(e => e.ID, e => e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex=3,AreaIndex=2,NewAreaIndex=3,Projects=Team,ID="2201",Values=new[]{"市","县","项目名称","项目规模","新增耕地面积"}});
        }
    }
}