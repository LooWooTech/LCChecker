using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    [Table("seprojects")]
    public class SecondProject:SecondProjectBase
    {               
        /// <summary>
        /// 总规模
        /// </summary>
        public double? Area { get; set; }
        /// <summary>
        /// 新增耕地面积
        /// </summary>
        public double? NewArea { get; set; }
        public double? SurplusHookArea { get; set; }
        public double? TrueHookArea { get; set; }
    }
}