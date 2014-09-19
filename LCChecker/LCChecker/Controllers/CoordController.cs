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
        public ActionResult CoordProjects(NullableFilter result = NullableFilter.All, int page = 1)
        {
            var filter = new ProjectFileter
            {
                City = CurrentUser.City,
                Result = result,
                Visible = true,
                Page = new Page(page)
            };
            ViewBag.List = ProjectHelper.GetCoordProjects(filter);
            ViewBag.Page = filter.Page;
            return View();
        }


        public ActionResult UploadCoordProjects(UploadFileType type)
        {
            throw new NotImplementedException();
        }
    }
}