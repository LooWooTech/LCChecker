using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class Mistake
    {
        public int Id { get; set; }
        public string Error { get; set; }//违反的规则
        public int row { get; set; }//第几行
        public int cell { get; set; }//第几列
    }
}