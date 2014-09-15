using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    [Table("reports")]
    public class Report
    {
        [Key]
        [MaxLength(55)]
        public string ID { get; set; }
        /// <summary>
        /// 市
        /// </summary>
        [Column("CityID", TypeName = "INT")]
        public City City { get; set; }


        /// <summary>
        /// 
        /// </summary>
        [Column("Type",TypeName = "INT")]
        public ReportType Type { get; set; }



        /// <summary>
        /// 验证标志位
        /// </summary>
        public bool? Result { get; set; }
    }
}