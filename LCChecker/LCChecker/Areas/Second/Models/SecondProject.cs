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
    public class SecondProject
    {
        public SecondProject() {
            UpdateTime = DateTime.Now;
        }
        [Key]
        [MaxLength(55)]
        public string ID { get; set; }
        [Column(TypeName="int")]
        public City City { get; set; }
        [MaxLength(255)]
        public string Name { get; set; }
        public bool? Result { get; set; }
        public string County { get; set; }
        [MaxLength(1023)]
        public string Note { get; set; }
        public DateTime UpdateTime { get; set; }
        /// <summary>
        /// 市级自查是否存在疑问
        /// </summary>
        public bool IsHasDoubt { get; set; }
        /// <summary>
        /// 属于申请删除项目
        /// </summary>
        public bool IsApplyDelete { get; set; }
        
        /// <summary>
        /// 备案信息存在错误项目
        /// </summary>
        public bool IsHasError { get; set; }
        /// <summary>
        /// 属于可用于占补平衡面积核减
        /// </summary>
        public bool IsDescrease { get; set; }
        /// <summary>
        /// 属于已挂钩使用需要接触挂钩关系可用于占补平衡
        /// </summary>
        public bool IsRelieve { get; set; }
        /// <summary>
        /// 总规模
        /// </summary>
        public double? Area { get; set; }
        /// <summary>
        /// 新增耕地面积
        /// </summary>
        public double? NewArea { get; set; }
    }
}