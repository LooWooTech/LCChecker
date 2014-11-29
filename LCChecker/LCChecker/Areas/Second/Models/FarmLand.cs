using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    [Table("farmland")]
    public class FarmLand
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }
        public string ProjectID { get; set; }
        [Column(TypeName="int")]
        public LandType Type { get; set; }
        public double Area { get; set; }
        [Column(TypeName="int")]
        public Degree Degree { get; set; }
    }

    public enum LandType { 
        [Description("水田")]
        Paddy=1,
        [Description("旱地")]
        Dry=2
    }

    public enum Degree { 
        一等=1,
        二等=2,
        三等=3,
        四等=4,
        五等=5,
        六等=6,
        七等=7,
        八等=8,
        九等=9,
        十等=10,
        十一等=11,
        十二等=12,
        十三等=13,
        十四等=14,
        十五等=15
    }
}