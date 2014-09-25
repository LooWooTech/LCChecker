using LCChecker.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    public class SpecialData:IRowRule
    {
        public int ColumnIndex { get; set; }//检验的列
        public int ColumnIndex2 { get; set; }
        public string Value { get; set; }
        public int IDIndex { get; set; }//项目编号所在的列
        public Dictionary<string, Index2> ProjectData { get; set; }
        public string ID { get; set; }

        public string Name {
            get{
                return string.Format("规则{0}：{1}与项目不符", ID,Value);
            }  
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value=row.GetCell(IDIndex+xoffset,MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if(string.IsNullOrEmpty(value))
                return false;
            if(!ProjectData.ContainsKey(value))
                return false;
            var item =ProjectData[value];
            var cell = row.GetCell(ColumnIndex + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            switch (Value)
            {
                case "耕地质量等别":
                    value = cell.ToString().Trim();
                    if (value != item.Grade)
                        return false;
                    break;
                case "已与建设项目预挂钩应核销占补平衡指标":
                    var cell2 = row.GetCell(ColumnIndex2 + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    double val=.0;
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
                    else {
                        var value2 = cell.ToString().Trim();
                        double.TryParse(value2, out val);
                    }
                    double val2;
                    if (cell2.CellType == CellType.Numeric || cell2.CellType == CellType.Formula)
                    {
                        try
                        {
                            val2 = cell2.NumericCellValue;
                        }
                        catch
                        {
                            val2 = .0;
                        }
                    }
                    else {
                        var value2 = cell2.ToString().Trim();
                        double.TryParse(value2, out val2);
                    }
                    double sum = val + val2;
                    if (Math.Abs(sum - item.Indicators) > 0.0001)
                        return false;
                    break;
                //case "新增耕地":
                    //value = cell.ToString().Trim();
                    //var Land = value.GetLand();
                    //if (!item.Land.Compare(Land))
                      //  return false;
                    //break;
                default: break;
            }
            return true;
        }
    }
}