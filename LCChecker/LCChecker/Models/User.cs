using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class User
    {
        [Key]
        [DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        [Column("Username")]
        [MaxLength(55)]
        public string Username { get; set; }

        [MaxLength(55)]
        public string Password { get; set; }

        public bool Flag { get; set; }

        [Column("CityID", TypeName = "INT")]
        public City City { get; set; }
    }
}