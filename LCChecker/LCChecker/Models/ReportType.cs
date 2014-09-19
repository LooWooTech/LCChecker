using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public enum ReportType
    {
        [Description("重点复核确认项目申请删除项目清单")]
        附表4 = 4,

        [Description("重点复核确认项目备案信息错误项目清单")]
        附表5 = 5,

        [Description("重点复核确认项目耕地质量等别修改项目清单")]
        附表7 = 7,

        [Description("重点复核确认项目占补平衡指标核减项目清单")]
        附表8 = 8,

        [Description("重点复核确认项目新增耕地二级地类与耕地质量等别确认表")]
        附表9 = 9
    }
}