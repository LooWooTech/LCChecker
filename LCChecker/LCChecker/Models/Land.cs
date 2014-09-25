using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class Land
    {
        /// <summary>
        /// 水田
        /// </summary>
        public double Paddy { get; set; }

        /// <summary>
        /// 水浇地
        /// </summary>
        public double Irrigated { get; set; }

        /// <summary>
        /// 旱地
        /// </summary>
        public double Dry { get; set; }


        public bool Compare(Land Data)
        {
            return (Math.Abs(this.Paddy - Data.Paddy) < 0.0001) && (Math.Abs(this.Irrigated - Data.Irrigated) < 0.0001) &&(Math.Abs(this.Dry - Data.Dry) < 0.0001);
        }
    }
}