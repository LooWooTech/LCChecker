using LCChecker.Models;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LCChecker.Areas.Second.Models
{
    public class SecondProjectBase
    {
        public SecondProjectBase() {
            UpdateTime = DateTime.Now;
        }
        
        [Column(TypeName = "int")]
        public City City { get; set; }
        [MaxLength(255)]
        public string Name { get; set; }
        
        public string County { get; set; }
        [MaxLength(1023)]
        public string Note { get; set; }
        public DateTime UpdateTime { get; set; }
       
        /// <summary>
        /// 属于申请删除项目
        /// </summary>
        public bool IsApplyDelete { get; set; }

        /// <summary>
        /// 备案信息存在错误项目
        /// </summary>
        public bool IsHasError { get; set; }
       

    }
}