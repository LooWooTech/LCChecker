﻿using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    internal class NoLessThanRowRule:IRowRule
    {
        public int Column1Index { get; set; }
        public int Column2Index { get; set; }
        public string ID { get; set; }
        public string Name {
            get { return string.Format("规则{0}：第{1}栏大于等于第{2}栏",ID, Column1Index + 1, Column2Index + 1); }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1 = row.GetCell(xoffset + Column1Index, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            var value2 = row.GetCell(xoffset + Column2Index, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            return SumRowRule.GetValue(value1) >= SumRowRule.GetValue(value2);
            //double val1, val2;
            
            //return double.TryParse(value1, out val1) && double.TryParse(value2, out val2) && val1 >= val2;

        }
    }
}