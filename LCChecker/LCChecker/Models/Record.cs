using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    [Table("records")]
    public class Record
    {
        [Key]
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        /// <summary>
        /// 项目编号
        /// </summary>
        [MaxLength(55)]
        public string ProjectID { get; set; }

        /// <summary>
        /// 报部表格
        /// </summary>
        [Column("Type",TypeName="INT")]
        public ReportType Type { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [Column("CityID",TypeName="INT")]
        public City City { get; set; }

        public bool IsError { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        [MaxLength(1023)]
        public string Note { get; set; }


    }
}