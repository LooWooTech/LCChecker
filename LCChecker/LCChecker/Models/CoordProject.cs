using LCChecker.Areas.Second.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    /// <summary>
    /// 项目坐标
    /// </summary>
    [Table("coord_projects")]
    public class CoordProject:CoordProjectBase
    {
        public CoordProject() {
            UpdateTime = DateTime.Now;
        }
    }
}