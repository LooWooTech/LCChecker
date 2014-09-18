using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    public class DoubleEqual : IRowRule
    {

        public int ColumnIndex { get; set; }
        public double data { get; set; }
        public string Name
        {
            get
            {
                return string.Format("第{0}栏填写的内容为{1}", ColumnIndex + 1, data);
            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex+xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (string.IsNullOrEmpty(value))
                return false;
            double a = double.Parse(value);
            return Math.Abs(a - data) < 0.0001;
            //if (data != a)
            //    return false;
            //return true;
        }
    }
}