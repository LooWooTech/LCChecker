using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LooWoo.Land.LandControlChecker.Rules
{
    internal class UniqueValueRowRule:IRowRule
    {
        public int ColumnIndex { get; set; }
        public IRowRule Rule { get; set; }
        public string Keyword { get; set; }
        private Dictionary<string, int> nameDict = new Dictionary<string, int>(); 
       
        public string Name 
        {
            get { return "项目名称包含“综合整治”需复核是否重复备案需删除"; }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1 = row.GetCell(xoffset + ColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();

            if (!string.IsNullOrEmpty(Keyword) && value1.Contains(Keyword))
            {
                if (nameDict.ContainsKey(value1))
                {
                    nameDict[value1] = nameDict[value1] + 1;
                    return false;
                }
                else
                {
                    nameDict[value1] = 1;
                    return false;
                }
            }
            return true;
        }
    }
}
