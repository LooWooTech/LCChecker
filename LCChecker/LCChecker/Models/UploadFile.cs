﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    [Table("uploadfiles")]
    public class UploadFile
    {
        public UploadFile()
        {
            CreateTime = DateTime.Now;
        }

        [Key]
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        [Column("CityID")]
        public City City { get; set; }

        [MaxLength(127)]
        public string FileName { get; set; }

        public DateTime CreateTime { get; set; }

        [MaxLength(127)]
        public string SavePath { get; set; }

        /// <summary>
        /// 处理状态
        /// </summary>
        [Column("State", TypeName = "INT")]
        public UploadFileProceedState State { get; set; }

        [MaxLength(1023)]
        public string ProcessMessage { get; set; }

        /// <summary>
        /// 文件类型
        /// </summary>
        [Column("Type", TypeName = "INT")]
        public UploadFileType Type { get; set; }


    }

    /// <summary>
    /// 上传文件类型
    /// </summary>
    public enum UploadFileType
    {
        自查表 = 0,
        附表四 = 4,
        附表五 = 5,
        附表七 = 7,
        附表八 = 8,
        附表九 = 9,
        项目坐标 = 10,
        新增耕地坐标 = 11
    }

    public enum UploadFileProceedState
    {
        UnProceed,
        Proceeded,
        Error
    }
}