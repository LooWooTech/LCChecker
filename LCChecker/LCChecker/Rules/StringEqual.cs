using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    public class StringEqual : IRowRule
    {

        public int ColumnIndex { get; set; }
        public string Data { get; set; }
        public string Name
        {
            get
            {
                return string.Format("第{0}栏填写的内容为{1}", ColumnIndex+1,Data);
            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex+xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (string.IsNullOrEmpty(value))
                return false;
            if (value != Data)
                return false;
            return true;
        }
    }
}