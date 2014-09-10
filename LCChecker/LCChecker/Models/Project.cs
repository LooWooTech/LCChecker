using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class Project
    {
        [Key]
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        /// <summary>
        /// 项目编号
        /// </summary>
        [Required]
        public string NO { get; set; }

        [Column("CityID", TypeName = "INT")]
        public City City { get; set; }

        /// <summary>
        /// 检查结果
        /// </summary>
        public bool? Result { get; set; }

        /// <summary>
        /// 备注
        /// </summary>
        public string Note { get; set; }
    }

}