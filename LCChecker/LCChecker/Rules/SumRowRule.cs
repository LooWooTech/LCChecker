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
        public string Name {
            get {
                var sb = new StringBuilder(string.Format("第{0}栏等于{1}栏", SumColumnIndex + 1, ColumnIndices[0] + 1));
                for (var i = 1; i < ColumnIndices.Length; i++)
                {
                    sb.AppendFormat("+{0}栏", ColumnIndices[i] + 1);
                }
                return sb.ToString();
            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1 = row.GetCell(xoffset + SumColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
            double sum = 0.0;
            double.TryParse(value1, out sum);

            var sum2 = 0.0;
            double val;
            var list = new List<double>();
            for (var i = 0; i < ColumnIndices.Length; i++)
            {
                var value2 = row.GetCell(xoffset + ColumnIndices[i], MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
                val = 0;
                double.TryParse(value2, out val);

                sum2 += val;
            }

            return Math.Abs(sum2 - sum) < 0.00001;
        }
    }
}