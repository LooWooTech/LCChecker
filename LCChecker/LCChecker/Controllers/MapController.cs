using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using LCChecker.Models;

namespace LCChecker.Controllers
{
    public partial class UserController
    {
        public ActionResult Map()
        {
            return View();
        }

        public ActionResult UpdateMapData()
        {
            throw new NotImplementedException();
        }
    }
}