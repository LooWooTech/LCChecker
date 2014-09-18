using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    internal class CellEmptyRowRule:IRowRule
    {
        public int ColumnIndex { get; set; }
        public bool isEmpty { get; set; }
        public bool isNumeric { get; set; }
        public string Name {
            get {
                if (isEmpty)
                {
                    return string.Format(isNumeric ? "第{0}栏无面积" : "第{0}栏无填写", ColumnIndex + 1);
                }
                else {
                    return string.Format(isNumeric ? "第{0}栏有面积" : "第{0}栏填写", ColumnIndex + 1);
                }

            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1=row.GetCell(xoffset+ColumnIndex,MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
            if (isEmpty && string.IsNullOrEmpty(value1))// isEmpty :true  里面空的  1）无填写
                return true;

            double sum;
            var ret = double.TryParse(value1, out sum);

            if (isNumeric)//  面积问题
            {
                if (ret)
                {
                    if (isEmpty) return Math.Abs(sum) < double.Epsilon;
                    return Math.Abs(sum) >= double.Epsilon;
                }
                return false;
                
            }
            else
            {
                return !(isEmpty ^ string.IsNullOrEmpty(value1));
            }
        }
    }
}