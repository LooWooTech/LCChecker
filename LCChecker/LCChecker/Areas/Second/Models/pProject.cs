using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace LCChecker.Areas.Second.Models
{
    [Table("pproject")]
    public class pProject:SecondProjectBase
    {
        [Key]
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }
        public string Key { get; set; }
        public bool IsRight { get; set; }
    }
}