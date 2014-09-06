using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class RuleInfo
    {
        private static int count = 0;

        public RuleInfo()
        {
            Id = count;
          //  Enabled = true;
            count++;
        }


        public int Id { get; set; }
        //public int SheetIndex { get; set; }//表格号
        //public int CheckSheetColumnIndex { get; set; }//行号
        public IRowRule Rule { get; set; }//验证rule
        //public bool Enabled { get; set; }//验证成功与否
    }
}