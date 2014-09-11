using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace LCChecker.Rules
{
    public class Special:IRowRule 
    {
        public int ColumnIndex { get; set; }

        public string Name {
            get {
                return string.Format("第{0}栏应填写地类同时填写面积（面积不能为负）,多地类使用逗号",ColumnIndex+1);
            }
        }
        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            string[] team = value.Split(',');
            foreach(string item in team)
            {
                Regex r=new Regex(@"-?[0-9]");
                string a = r.Match(item).ToString();
                int position = item.IndexOf(a);
                if (position == 0)
                    return false;
                string b = item.Substring(position);
                double c;
                double.TryParse(b, out c);
                if (c < double.Epsilon)
                    return false;
            }
            return true;
        }
    }
}