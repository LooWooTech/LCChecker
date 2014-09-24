using LCChecker.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    public class OnlyProject:IRowRule
    {
        public int ColumnIndex { get; set; }
        public Dictionary<string, Project> Projects { get; set; }

        public string Name { get {
            return string.Format("第{0}栏项目编号不符",ColumnIndex+1);
        } }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (string.IsNullOrEmpty(value))
                return false;
            if (!Projects.ContainsKey(value))
                return false;
            var project = Projects[value];
            var division="浙江省,"+project.City.ToString()+","+project.County;
            value = row.GetCell(ColumnIndex + xoffset - 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (value != division)
                return false;
            value = row.GetCell(ColumnIndex + xoffset + 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (value != project.Name)
                return false;
            return true;
        }
    }
}