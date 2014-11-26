using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Areas.Second.Controllers
{
    public class HomeController : SecondController
    {
        //
        // GET: /Second/Home/

        public ActionResult Index()
        {
            return View();
        }

    }
}
