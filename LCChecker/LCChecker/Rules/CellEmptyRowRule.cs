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
            var cell = row.GetCell(xoffset + ColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);

            var ret = false;
            double sum;
            if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
            {
                ret = true;
                try
                {
                    sum = cell.NumericCellValue;
                }
                catch {
                    sum = .0;
                }
                
            }
            else
            {
                var value2 = cell.ToString();
                ret = double.TryParse(value2, out sum);
            }

            var value1=row.GetCell(xoffset+ColumnIndex,MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
            
            if (isEmpty && string.IsNullOrEmpty(value1))
                return true;

            if (isNumeric)//  面积问题
            {
                if (ret)
                {
                    return isEmpty ^ Math.Abs(sum) >= double.Epsilon; 
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