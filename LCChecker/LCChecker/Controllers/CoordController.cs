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
        public ActionResult Coord(int page = 1)
        {
            return View();
        }

        public ActionResult UpdateCoord(UploadFileType type)
        {
            throw new NotImplementedException();
        }
    }
}