using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    internal class AndRule:IRowRule
    {
        public IRowRule Rule1 { get; set; }
        public IRowRule Rule2 { get; set; }

        public string Name {
            get { return string.Format("{0},并且{1}", Rule1.Name, Rule2.Name); }
        }

        public bool Check(NPOI.SS.UserModel.IRow row,int xoffset=0)
        {
            return Rule1.Check(row,xoffset)&&Rule2.Check(row,xoffset);
        }
    }
}