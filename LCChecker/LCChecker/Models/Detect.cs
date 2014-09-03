using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    //各个区域提交情况
    public class Detect
    {
        public int Id { get; set; }
        public string region { get; set; }//区域
        public int sum { get; set; }//项目个数
        public int submit { get; set; }//提交次数
        public int Correct { get; set; }//其中符合要求  正确的个数
        public int totalScale { get; set; }//总规模
        public int AddArea { get; set; }//新增耕地面积
        public int available { get; set; }//可用于占补平衡面积
        public bool flag { get; set; }//提交的项目是否全部通过
    }
}