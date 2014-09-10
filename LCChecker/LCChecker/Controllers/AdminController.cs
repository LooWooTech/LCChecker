using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    [UserAuthorize]
    public class AdminController : ControllerBase
    {
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 上传总表数据
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects()
        {
            return View();
        }
    }
}
