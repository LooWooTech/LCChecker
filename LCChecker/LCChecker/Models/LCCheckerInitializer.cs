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
            context.USER.Add(new User { Username = "admin", City = City.浙江省, Password = "admin", Flag = true });
            context.USER.Add(new User { Username = "quzhou", City = City.衢州市, Password = "0570", Flag = false });
            context.USER.Add(new User { Username = "hangzhou", City = City.杭州市, Password = "0571", Flag = false });
            context.USER.Add(new User { Username = "huzhou", City = City.湖州市, Password = "0572", Flag = false });
            context.USER.Add(new User { Username = "jiaxin", City = City.嘉兴市, Password = "0573", Flag = false });
            context.USER.Add(new User { Username = "ningbo", City = City.宁波市, Password = "0574", Flag = false });
            context.USER.Add(new User { Username = "shaoxing", City = City.绍兴市, Password = "0575", Flag = false });
            context.USER.Add(new User { Username = "taizhou", City = City.台州市, Password = "0576", Flag = false });
            context.USER.Add(new User { Username = "wenzhou", City = City.温州市, Password = "0577", Flag = false });
            context.USER.Add(new User { Username = "lishui", City = City.丽水市, Password = "0578", Flag = false });
            context.USER.Add(new User { Username = "jinhua", City = City.金华市, Password = "0579", Flag = false });
            context.USER.Add(new User { Username = "zhoushan", City = City.舟山市, Password = "0580", Flag = false });


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