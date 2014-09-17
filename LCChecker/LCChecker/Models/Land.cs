using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class Land
    {
        public double Paddy { get; set; }

        public double Irrigated { get; set; }

        public double Dry { get; set; }


        public bool Compare(Land Data)
        {
            return (Math.Abs(this.Paddy - Data.Paddy) < double.Epsilon) && (Math.Abs(this.Irrigated - Data.Irrigated) < double.Epsilon) &&(Math.Abs(this.Dry - Data.Dry) < double.Epsilon);
        }
    }
}