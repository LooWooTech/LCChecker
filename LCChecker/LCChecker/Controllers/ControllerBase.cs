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

        protected ActionResult JsonSuccess(object data = null)
        {
            return Json(new { result = true, data }, JsonRequestBehavior.AllowGet);
        }

        protected ActionResult JsonError(string message = null)
        {
            return Json(new { result = false, message }, JsonRequestBehavior.AllowGet);
        }

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            ViewBag.CurrentUser = CurrentUser;
            base.OnActionExecuting(filterContext);
        }

        private Exception GetException(Exception ex)
        {
            var innerEx = ex.InnerException;
            if (innerEx != null)
            {
                return GetException(innerEx);
            }
            return ex;
        }

        protected override void OnException(ExceptionContext filterContext)
        {
            if (filterContext.ExceptionHandled) return;
            filterContext.ExceptionHandled = true;
            filterContext.HttpContext.Response.StatusCode = 500;
            ViewBag.Exception = GetException(filterContext.Exception);// filterContext.Exception;
            filterContext.Result = View("Error");
        }
    }
}