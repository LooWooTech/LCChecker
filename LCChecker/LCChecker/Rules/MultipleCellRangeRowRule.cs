using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace LCChecker.Rules
{
    internal class MultipleCellRangeRowRule:IRowRule
    {
        public int[] ColumnIndices { get; set; }
        public bool isEmpty { get; set; }
        public bool isNumeric { get; set; }
        public bool isAny { get; set; }
        public string Name {
            get {
                var sb = new StringBuilder();
                foreach (var column in ColumnIndices)
                {
                    if (sb.Length > 0)
                        sb.Append(",");
                    sb.AppendFormat("第{0}栏", column + 1);
                }
                if (isAny)
                    sb.Append("中至少有一栏");

                if (isEmpty)
                {
                    sb.Append(isNumeric ? "无面积" : "无填写");
                }
                else {
                    sb.Append(isNumeric ? "有面积" : "已填写");
                }
                return sb.ToString();
            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var rule = new CellEmptyRowRule()
            {
                isEmpty = this.isEmpty,
                isNumeric = this.isNumeric
            };
            if (isAny)
            {
                foreach (var index in ColumnIndices)
                {
                    rule.ColumnIndex = index;
                    if (rule.Check(row, xoffset))
                        return true;
                }
                return false;
            }
            else {
                foreach (var index in ColumnIndices)
                {
                    rule.ColumnIndex = index;
                    if (!rule.Check(row, xoffset))
                        return false;
                }
                return true;
            }
        }

    }
}