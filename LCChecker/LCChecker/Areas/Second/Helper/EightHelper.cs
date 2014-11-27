using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public class EightHelper
    {

    }

    public class HookedProject {
        public City City { get; set; }
        public string County { get; set; }
        public string ID { get; set; }
        public string Name { get; set; }
        public double NewArea { get; set; }
        public double Balance { get; set; }
        public double HookedBalance { get; set; }
        public double RelieveBalance { get; set; }
    }

    public class BuildProject {
        public City City { get; set; }
        public string County { get; set; }
        public string ID { get; set; }
        public string Name { get; set; }
        public double BuildBalance { get; set; }
    }

    public class ReHookProject {
        public City City { get; set; }
        public string County { get; set; }
        public string ID { get; set; }
        public string Name { get; set; }
        public double AvailBalance { get; set; }
        public double BeBalance { get; set; }
        public double BuildBalance { get; set; }
        public double AfBalance { get;set; }
    }
}