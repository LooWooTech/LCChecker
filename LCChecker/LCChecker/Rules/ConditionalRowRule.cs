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

        public string ID { get; set; }
        public string Name {
            get { return string.Format("规则{0}：{1} 时，{2}",ID,Condition.Name,Rule.Name);}
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            if (Condition.Check(row, xoffset))
                return Rule.Check(row, xoffset);
            return true;
        }
    }
}