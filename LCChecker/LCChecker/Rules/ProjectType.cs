using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace LCChecker.Rules
{
    public class ProjectType:IRowRule
    {
        public int Time { get; set; }

        public string[] Variety { get; set; }
        public string Name {
            get { 
                var sb = new StringBuilder();
                sb.AppendFormat("项目类型为{0}年前",Time);
                foreach(var item in Variety)
                {
                    sb.AppendFormat("土地{0}",item);
                }

                return sb.ToString();
            }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            bool flag=false;
            var value = row.GetCell(xoffset + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            string vtime = value.Substring(6, 4);
            int a = int.Parse(vtime);
            if (a > Time)
                return false;
            var value2 = row.GetCell(xoffset + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            foreach (var item in Variety)
            {
                if (value2.Contains(item))
                    flag = true;
            }
            return flag;
        }
    }
}