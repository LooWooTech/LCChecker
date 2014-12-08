using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    [Table("coord_newareaprojects")]
    public class CoordNewAreaProject:CoordProjectBase
    {
        public CoordNewAreaProject() {
            UpdateTime = DateTime.Now;
        }
    }
}