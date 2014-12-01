using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    public enum RuleKind
    {
        [Description("全部")]
        All=0,
        [Description("基本规则")]
        Basic=1,
        [Description("填写规则")]
        Write=2,
        [Description("一致性")]
        Consistency=3,
        [Description("数据规则")]
        Data=4,
        [Description("其他")]
        Other=5
    }
}