using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class LCCheckerInitializer:DropCreateDatabaseAlways<LCDbContext>
    {
        protected override void Seed(LCDbContext context)
        {
            context.USER.Add(new User { logName = "admin", name = "管理员", password = "admin", flag = true });
            context.USER.Add(new User { logName = "quzhou", name = "衢州市", password = "0570", flag = false });
            context.USER.Add(new User { logName = "hangzhou", name = "杭州市", password = "0571", flag = false });
            context.USER.Add(new User { logName = "huzhou", name = "湖州市", password = "0572", flag = false });
            context.USER.Add(new User { logName = "jiaxin", name = "嘉兴市", password = "0573", flag = false });
            context.USER.Add(new User { logName = "ningbo", name = "宁波市", password = "0574", flag = false });
            context.USER.Add(new User { logName = "shaoxing", name = "绍兴市", password = "0575", flag = false });
            context.USER.Add(new User { logName = "taizhou", name = "台州市", password = "0576", flag = false });
            context.USER.Add(new User { logName = "wenzhou", name = "温州市", password = "0577", flag = false });
            context.USER.Add(new User { logName = "lishui", name = "丽水市", password = "0578", flag = false });
            context.USER.Add(new User { logName = "jinhua", name = "金华市", password = "0579", flag = false });
            context.USER.Add(new User { logName = "zhoushan", name = "舟山市", password = "0580", flag = false });


            context.DETECT.Add(new Detect { region = "杭州市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "湖州市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "嘉兴市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "宁波市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "绍兴市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "台州市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "温州市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "丽水市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "金华市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "舟山市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
            context.DETECT.Add(new Detect { region = "衢州市", sum = 0, submit = 0, Correct = 0, totalScale = 0, AddArea = 0, available = 0, flag = false });
        }

    }
}