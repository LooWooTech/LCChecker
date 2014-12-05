using LCChecker.Areas.Second.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace LCChecker.Areas.Second.Rules
{
    public class OnlySecondProject:IRowRule
    {
        public int ColumnIndex { get; set; }
        public string[] Values { get; set; }
        public int AreaIndex { get; set; }
        public int NewAreaIndex { get; set; }
        public Dictionary<string, SecondProject> Projects { get; set; }
        public string ID { get; set; }

        public string Name
        {
            get
            {
                var sb = new StringBuilder(string.Format("规则{0}:'{1}'", ID, Values[0]));
                for (var i = 1; i < Values.Length; i++)
                {
                    sb.AppendFormat("、'{0}'", Values[i]);
                }
                sb.AppendFormat("与重点复核确认项目清单中保持一致");
                return sb.ToString();
            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (string.IsNullOrEmpty(value))
                return false;
            if (!Projects.ContainsKey(value))
                return false;
            var project = Projects[value];
            for (var i = 1; i < Values.Length; i++)
            {
                var item = Values[i];
                switch (item)
                {
                    case "行政区":
                        var division = "浙江省," + project.City.ToString() + "," + project.County;
                        value = row.GetCell(ColumnIndex + xoffset - 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        value = value.Replace(" ", "").Replace("，", ",");
                        if (value != division)
                            return false;
                        break;
                    case "项目名称":
                        value = row.GetCell(ColumnIndex + xoffset + 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        value = value.Replace(" ", "");
                        var str = project.Name.Replace(" ", "");
                        if (value != str)
                            return false;
                        break;
                    case "市":
                        value = row.GetCell(ColumnIndex + xoffset - 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        value = value.Replace(" ", "");
                        if (value != project.City.ToString())
                            return false;
                        break;
                    case "县":
                        value = row.GetCell(ColumnIndex + xoffset - 1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        value = value.Replace(" ", "");
                        if (value != project.County)
                            return false;
                        break;
                    case "新增耕地面积":
                        var cell = row.GetCell(ColumnIndex + xoffset+NewAreaIndex , MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        double val = .0;
                        if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                        {
                            try
                            {
                                val = cell.NumericCellValue;
                            }
                            catch
                            {
                                val = .0;
                            }
                        }
                        var CurrentVal = .0;
                        if (project.NewArea.HasValue)
                            CurrentVal = project.NewArea.Value;
                        if (Math.Abs(val - CurrentVal) > 0.0001)
                            return false;
                        break;
                    case "项目规模":
                        var cells = row.GetCell(ColumnIndex + xoffset + AreaIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        double vals = .0;
                        if (cells.CellType == CellType.Numeric || cells.CellType == CellType.Formula) {
                            try
                            {
                                vals = cells.NumericCellValue;
                            }
                            catch {
                                vals = .0;
                            }
                        }
                        var CurrentVals = .0;
                        if (project.Area.HasValue)
                            CurrentVals = project.Area.Value;
                        if (Math.Abs(vals - CurrentVals) > 0.0001)
                            return false;
                        break;
                    case "剩余可用于占补平衡面积":
                        var surpluecell = row.GetCell(ColumnIndex + xoffset + 5, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        double sval = 0.0;
                        if (surpluecell.CellType == CellType.Numeric || surpluecell.CellType == CellType.Formula) {
                            try
                            {
                                sval = surpluecell.NumericCellValue;
                            }
                            catch {
                                sval = 0.0;
                            }
                        }
                        var SurCurrentVal = 0.0;
                        if (project.SurplusHookArea.HasValue) {
                            SurCurrentVal = project.SurplusHookArea.Value;
                        }
                        if (Math.Abs(sval - SurCurrentVal) > 0.0001)
                            return false;
                        break;
                    case "实际可用于占补平衡面积":
                        var hookCell = row.GetCell(ColumnIndex + xoffset + 4, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        double hookval = 0.0;
                        if (hookCell.CellType == CellType.Numeric || hookCell.CellType == CellType.Formula) {
                            try
                            {
                                hookval = hookCell.NumericCellValue;
                            }
                            catch {
                                hookval=0.0;
                            }
                        }
                         var hookCurrent = 0.0;
                         if (project.TrueHookArea.HasValue) {
                             hookCurrent = project.TrueHookArea.Value;
                         }
                         if (Math.Abs(hookval - hookCurrent) > 0.0001)
                             return false;
                         break;
                    default: break;
                }
            }
            return true;
        }
    }
}