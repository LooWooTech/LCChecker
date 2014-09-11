using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    public class HomeController : ControllerBase
    {
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Login(string username, string password)
        {
            var user = db.USER.FirstOrDefault(e => e.name.ToLower() == username.ToLower());
            if (user == null)
            {
                throw new ArgumentException("用户不存在");
            }

            if (user.password != password)
            {
                throw new ArgumentException("密码不正确");
            }

            CurrentUser = user;


            return RedirectToAction("Index", user.flag ? "Admin" : "User");
        }


        public ActionResult Logout()
        {
            CurrentUser = null;
            return RedirectToAction("Index");
        }

    }
}
