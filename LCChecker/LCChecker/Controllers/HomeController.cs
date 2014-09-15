using System;
using System.Collections.Generic;
using System.IO;
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
            if (CurrentUser != null)
            {
                if (CurrentUser.Flag)
                {
                    return RedirectToAction("Index", "Admin");
                }
                else
                {
                    return RedirectToAction("Index", "User");
                }
            }

            return View();
        }

        [HttpPost]
        public ActionResult Login(string username, string password)
        {
            var user = db.Users.FirstOrDefault(e => e.Username.ToLower() == username.ToLower());
            if (user == null)
            {
                throw new ArgumentException("用户不存在");
            }

            if (user.Password != password)
            {
                throw new ArgumentException("密码不正确");
            }

            CurrentUser = user;


            return RedirectToAction("Index", user.Flag ? "Admin" : "User");
        }


        public ActionResult Logout()
        {
            CurrentUser = null;
            return RedirectToAction("Index");
        }

    }
}
