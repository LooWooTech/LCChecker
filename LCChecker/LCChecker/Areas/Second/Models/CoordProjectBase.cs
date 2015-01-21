﻿using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    public class CoordProjectBase
    {
        /// <summary>
        /// 项目编号
        /// </summary>
        [Key]
        [MaxLength(55)]
        public string ID { get; set; }

        [Column("CityID", TypeName = "INT")]
        public City City { get; set; }

        [MaxLength(255)]
        public string Name { get; set; }

        /// <summary>
        /// 检查结果
        /// </summary>
        public bool? Result { get; set; }

        /// <summary>
        /// 所在县
        /// </summary>
        public string County { get; set; }

        /// <summary>
        /// 备注
        /// </summary>
        [MaxLength(1023)]
        public string Note { get; set; }

        [MaxLength(1023)]
        public string Error { get; set; }

        public DateTime UpdateTime { get; set; }

        public bool Visible { get; set; }
        public bool Exception { get; set; }
    }

}