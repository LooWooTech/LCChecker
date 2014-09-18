﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    //ProjectName,ProjectNo, 县、市、是否申请删除、是否确认无误  是否是指标核减、是否需要修改
    [Table("projects")]
    public class Project
    {
        public Project()
        {
            UpdateTime = DateTime.Now;
        }
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

        public DateTime UpdateTime { get; set; }
    }

    //public enum ProjectType
    //{
    //    确认存疑 = 1,
    //    确认无误 = 2,
    //    申请删除 = 3
    //}
}