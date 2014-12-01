using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace LCChecker.Rules
{
    internal class SumRowRule:IRowRule
    {
        public int SumColumnIndex { get; set; }
        public int[] ColumnIndices { get; set; }
        public string ID { get; set; }
        public string Name {
            get {
                var sb = new StringBuilder(string.Format("规则{0}：第{1}栏等于{2}栏",ID, SumColumnIndex + 1, ColumnIndices[0] + 1));
                for (var i = 1; i < ColumnIndices.Length; i++)
                {
                    sb.AppendFormat("+{0}栏", ColumnIndices[i] + 1);
                }
                return sb.ToString();
            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var sum = GetValue(row.GetCell(xoffset + SumColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK));
            
            var sum2 = 0.0;
            
            for (var i = 0; i < ColumnIndices.Length; i++)
            {
                var val = GetValue(row.GetCell(xoffset + ColumnIndices[i], MissingCellPolicy.CREATE_NULL_AS_BLANK));
                sum2 += val;
            }

            return Math.Abs(sum2 - sum) < 0.0001;
        }

        public static double GetValue(ICell cell)
        {
            if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
            {
                try
                {
                    return cell.NumericCellValue;
                }
                catch
                {
                    return .0;
                }
            }
            else
            {
                var value2 = cell.ToString();
                double val;
                double.TryParse(value2, out val);
                return val;
            }
        }
    }
}