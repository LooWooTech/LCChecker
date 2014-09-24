using LCChecker.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace LCChecker.Rules
{
    public class OnlyProject:IRowRule
    {
        public int ColumnIndex { get; set; }
        public string[] Values { get; set; }
        public Dictionary<string, Project> Projects { get; set; }

        public string Name { get {
            var sb = new StringBuilder(string.Format("{0}", Values[0]));
            for (var i = 1; i < Values.Length; i++)
            {
                sb.AppendFormat("或{0}", Values[i]);
            }
            sb.AppendFormat("不符");
            return sb.ToString();
        } }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (string.IsNullOrEmpty(value))
                return false;
            if (!Projects.ContainsKey(value))
                return false;
            var project = Projects[value];
            for (var i = 1; i < Values.Length;i++ )
            {
                var item = Values[i];
                switch (item)
                {
                    case "行政区":
                        var division = "浙江省," + project.City.ToString() + "," + project.County;
                        value = row.GetCell(ColumnIndex + xoffset - 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (value != division)
                            return false;
                        break;
                    case "项目名称":
                        value = row.GetCell(ColumnIndex + xoffset + 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (value != project.Name)
                            return false;
                        break;
                    case "市":
                        value = row.GetCell(ColumnIndex + xoffset - 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (value != project.City.ToString())
                            return false;
                        break;
                    case "县":
                        value = row.GetCell(ColumnIndex + xoffset - 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (value != project.County)
                            return false;
                        break;
                    case "新增耕地面积":
                        var cell = row.GetCell(ColumnIndex + xoffset + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        double val=.0;
                        if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                        {
                            try
                            {
                                val = cell.NumericCellValue;
                            }
                            catch {
                                val = .0;
                            }
                        }
                        var CurrentVal=.0;
                        if (project.NewArea.HasValue)
                            CurrentVal = project.NewArea.Value;
                        if (Math.Abs(val - CurrentVal) > 0.0001)
                            return false;
                        break;
                        
                    default: break;
                }
            }
            return true;
        }
    }
}