using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    public enum SecondReportType
    {
        [Description("重点复核确认项目以外所有报部备案项目复核确认总表")]
        附表1=1,
        [Description("重点复核确认项目以外所有报部备案复核确认无问题项目清单")]
        附表2=2,
        [Description("重点复核确认项目以外所有报部备案复核确认申请删除项目清单")]
        附表3=3,
        [Description("重点复核确认项目以外所有报部备案复核确认备案信息错误项目清单")]
        附表4=4,
        [Description("重点复核确认项目以外所有报部备案复核确认耕地质量等别修改项目清单")]
        附表6=6,
        [Description("重点复核确认项目以外所有报部备案复核确认可用于占补平衡面积核减项目清单")]
        附表7=7,
        [Description("重点复核确认项目以外所有报部备案复核确认补充耕地项目与建设用地项目解挂重挂对应项目关系确认表")]
        附表8=8,
        [Description("重点复核确认项目以外所有报部备案复核确认项目新增耕地二级地类与耕地质量等别确认表")]
        附表9=9
    }
}