using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class User
    {
        public int id { get; set; }
        public string logName { get; set; }
        public string name { get; set; }
        public string password { get; set; }
        public bool flag { get; set; }

        [Column("CityID", TypeName = "INT")]
        public City City { get; set; }
    }
}