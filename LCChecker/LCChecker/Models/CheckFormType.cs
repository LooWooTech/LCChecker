using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public enum CheckFormType
    {
        重点复核确认无问题项目清单 = 2,
        重点项目复核确认总表 = 3,
        重点复核确认项目申请删除项目清单 = 4,
        重点复核确认项目备案信息错误项目清单 = 5,
        重点复核确认项目设计二调新增耕地项目清单 = 6,
        重点复核确认项目耕地质量等别修改项目清单 = 7,
        重点复核确认项目占补平衡指标核减项目清单 = 8,
        重点复核确认项目新增耕地二级地类与耕地质量等别确认表 = 9
    }
}