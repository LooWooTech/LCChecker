using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    public class Arrange:IRowRule
    {
        public int Time { get; set; }
        public string Variety { get; set; }

        public string Name {
            get {
                return string.Format("{0}年以后验收土地{1}项目", Time, Variety);
            }
        }
        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(xoffset + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            string vtime = value.Substring(6,4);
            int a = int.Parse(vtime);
            if (a <= Time)
                return false;
            var value2 = row.GetCell(xoffset + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (value2.Contains(Variety) && value2.Contains("宅基地") == false)
                return true;
            return false;
         
        }
    }
}