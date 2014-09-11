using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class UploadFile
    {
        [Key]
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        public City City { get; set; }

        [MaxLength(55)]
        public string FileName { get; set; }

        public DateTime CreateTime { get; set; }

        [MaxLength(55)]
        public string SavePath { get; set; }
    }
}