using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class SubRecord
    {
        public int id { get; set; }
        public string Format { get; set; }//文件格式
        public string name { get; set; }//文件名
        public int submits { get; set; }//该区域第几次提交表格
        public int regionId { get; set; }//区域id
        public string FileName {//文件全名
            get {
                return "NO"+submits + this.Format;
            }
        }
    }
}