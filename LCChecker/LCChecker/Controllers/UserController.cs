using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    [UserAuthorize]
    public class UserController : ControllerBase
    {
        private List<Project> GetProjects()
        {
            var query = db.Projects.Where(e => e.City == CurrentUser.City);

            return query.ToList();
        }


        public ActionResult Index(bool? result, int page = 1)
        {

            ViewBag.Projects = GetProjects();

            return View();
        }

        /*用户上传文件*/
        [HttpPost]
        public ActionResult FileUpload(FormCollection form)
        {
            return View();
        }
    }
}
