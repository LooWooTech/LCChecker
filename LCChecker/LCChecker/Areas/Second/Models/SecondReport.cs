using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    [Table("sereports")]
    public class SecondReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }
        [Column(TypeName="int")]
        public City City { get; set; }
        [Column(TypeName="int")]
        public SecondReportType Type { get; set; }
        public bool? Result { get; set; }
        [MaxLength(1023)]
        public string Note { get; set; }


    }
}