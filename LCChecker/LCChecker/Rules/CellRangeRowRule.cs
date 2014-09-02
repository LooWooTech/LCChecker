using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace LCChecker.Rules
{
    internal class CellRangeRowRule:IRowRule
    {
        public int ColumnIndex { get; set; }
        public string[] Values { get; set; }
        public string Name {
            get {
                var sb = new StringBuilder(string.Format("第{0}栏是否全部选‘{1}’", ColumnIndex + 1, Values[0]));
                for (var i = 0; i < Values.Length; i++)
                {
                    sb.AppendFormat("或“{0}”",Values[i]+1);
                }
                return sb.ToString();
            }
        }
        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1 = row.GetCell(xoffset + ColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            foreach (var star in Values)
            {
                if (star == value1)
                    return true;
            }
            return false;
        }

    }
}