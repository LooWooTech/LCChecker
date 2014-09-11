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
        /// <summary>
        /// 项目编号
        /// </summary>
        [Key]
        public string ID { get; set; }

        [Column("CityID", TypeName = "INT")]
        public City City { get; set; }

        public string Name { get; set; }

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