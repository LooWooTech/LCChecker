using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace LCChecker.Areas.Second.Rules
{
    public class Less:IRowRule
    {
        public int[] ColumnIndex { get; set; }
        public string Value { get; set; }
        public string ID { get; set; }
        public string Name { 
            get {
                var sb = new StringBuilder(string.Format("规则{0}:", ID));
                for (var i = 0; i < ColumnIndex.Length; i++) {
                    sb.AppendFormat("{0}栏 ", ColumnIndex[i]+1);
                }
                sb.AppendFormat("中至少有一栏填写{0}",Value);
                return sb.ToString();
            }
        }
        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            for (var i = 0; i < ColumnIndex.Length; i++) {
                var val = row.GetCell(ColumnIndex[i] + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(val))
                    continue;
                if (val == Value)
                    return true;
            }
            return false;
        }
    }
}