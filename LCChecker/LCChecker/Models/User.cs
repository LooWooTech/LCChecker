using System;
using System.Collections.Generic;
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
    }
}