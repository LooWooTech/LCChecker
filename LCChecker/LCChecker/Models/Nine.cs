using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class Nine
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
        [Column("Formtype",TypeName = "INT")]
        public CheckFormType CheckFormType { get; set; }



        /// <summary>
        /// 验证标志位
        /// </summary>
        public bool? Result { get; set; }
    }
}