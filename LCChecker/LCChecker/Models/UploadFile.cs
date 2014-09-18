using System;
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
        [Key]
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        [Column("CityID")]
        public City City { get; set; }

        [MaxLength(55)]
        public string FileName { get; set; }

        public DateTime CreateTime { get; set; }

        [MaxLength(55)]
        public string SavePath { get; set; }

        /// <summary>
        /// 0是自查，1-9是报部，10及以上是坐标
        /// </summary>
        public int Type { get; set; }
    }
}