using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    public class ConditionalRowRule:IRowRule
    {
        public IRowRule Condition { get; set; }
        public IRowRule Rule { get; set; }

        public string Name {
            get { return string.Format("{0} 时，{1}",Condition.Name,Rule.Name);}
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            if (Condition.Check(row, xoffset))
                return Rule.Check(row, xoffset);
            return true;
        }
    }
}