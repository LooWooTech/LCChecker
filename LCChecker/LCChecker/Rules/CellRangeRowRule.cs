using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace LCChecker.Rules
{
    internal class CellRangeRowRule:IRowRule
    {
        public int ColumnIndex { get; set; }
        public string[] Values { get; set; }
        public string Name {
            get {
                var sb = new StringBuilder(string.Format("第{0}栏填写为：‘{1}’ ", ColumnIndex + 1, Values[0]));
                for (var i = 1; i < Values.Length; i++)
                {
                    sb.AppendFormat("或'{0}'",Values[i]);
                }
                return sb.ToString();
            }
        }
        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1 = row.GetCell(xoffset + ColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            foreach (var star in Values)
            {
                if (star == "经核实，项目由于___原因未实施或未终止实施，详细说明具体情况")
                {
                    return Regex.IsMatch(value1, @"^经核实[,，]项目由于([\w\W]+)原因未实施或未终止实施[,，]详细说明具体情况$");
                }
                if (star == value1)
                    return true;
            }
            return false;
        }

    }
}