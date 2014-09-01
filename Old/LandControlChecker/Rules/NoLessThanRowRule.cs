using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LooWoo.Land.LandControlChecker.Rules
{
    internal class NoLessThanRowRule:IRowRule
    {
        public int Column1Index { get; set; }
        public int Column2Index { get; set; }
        public string Name
        {
            get { return string.Format("第{0}栏大于等于第{1}栏", Column1Index + 1, Column2Index + 1); }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1 = row.GetCell(xoffset + Column1Index, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
            var value2 = row.GetCell(xoffset + Column1Index, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();

            double val1, val2;
            return double.TryParse(value1, out val1) && double.TryParse(value2, out val2) && val1 >= val2;
        }
    }
}
