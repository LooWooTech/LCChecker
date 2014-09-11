using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    public class ControllerBase : AsyncController
    {
        protected User CurrentUser
        {
            get
            {
                return AuthUtility.GetCurrentUser(HttpContext);
            }
            set
            {
                AuthUtility.SaveCurrentUser(HttpContext, value);
            }
        }

        private LCDbContext _db;
        protected LCDbContext db
        {
            get { return _db == null ? _db = new LCDbContext() : _db; }
        }

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            ViewBag.CurrentUser = CurrentUser;
            base.OnActionExecuting(filterContext);
        }
    }
}