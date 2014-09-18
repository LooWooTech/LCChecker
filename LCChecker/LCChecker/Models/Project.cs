﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
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

        //[Column("Type", TypeName = "INT")]
        //public ProjectType Type { get; set; }

        /// <summary>
        /// 备注
        /// </summary>
        [MaxLength(1023)]
        public string Note { get; set; }

        public DateTime UpdateTime { get; set; }

        /// <summary>
        /// 项目总规模（公顷）
        /// </summary>
        public double? Area { get; set; }

        /// <summary>
        /// 新增耕地面积
        /// </summary>
        public double? NewArea { get; set; }
    }

    public enum ProjectType
    {
        确认修改 = 1,
        确认删除 = 2
    }
}