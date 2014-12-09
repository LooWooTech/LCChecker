using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class Summary
    {
        public int TotalCount { get; set; }

        public City City { get; set; }

        public int SuccessCount { get; set; }

        public int ErrorCount { get; set; }
        public int ExceptionCount { get; set; }

        public int UnCheckCount { get { return TotalCount - SuccessCount - ErrorCount-ExceptionCount; } }
    }
}