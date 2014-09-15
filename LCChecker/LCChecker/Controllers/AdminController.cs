using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    [UserAuthorize]
    public class AdminController : ControllerBase
    {
        private void AddProjects(List<Project> list)
        {
            foreach (var item in list)
            {
                if (!db.Projects.Any(e => e.ID == item.ID))
                {
                    db.Projects.Add(item);
                }
                db.SaveChanges();
            }
        }

        public ActionResult Index()
        {
            ViewBag.Summary = db.Projects.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            return View();
        }

        public ActionResult Users()
        {
            ViewBag.List = db.Users.ToList();
            return View();
        }

        [HttpGet]
        public ActionResult AddUser()
        {
            return View();
        }
        [HttpPost]
        public ActionResult AddUser(User user)
        {
            if (string.IsNullOrEmpty(user.Username))
            {
                return JsonError("用户名没有填写");
            }

            if (string.IsNullOrEmpty(user.Password))
            {
                return JsonError("密码没有填写");
            }

            var rePwd = Request.Form["repassword"];
            if (string.IsNullOrEmpty(rePwd))
            {
                return JsonError("确认密码没有填写");
            }

            if (rePwd != user.Password)
            {
                return JsonError("密码输入不一致");
            }

            if (user.City == 0 && user.Flag == false)
            {
                return JsonError("非管理员用户请选择城市");
            }

            if (db.Users.Any(e => e.Username.ToLower() == user.Username.ToLower()))
            {
                return JsonError("用户名已被占用");
            }

            db.Users.Add(user);
            db.SaveChanges();

            return JsonSuccess();
        }

        public ActionResult DeleteUser(int id)
        {
            try
            {
                var entity = db.Users.FirstOrDefault(e => e.ID == id);
                db.Users.Remove(entity);
                db.SaveChanges();
                return JsonSuccess();
            }
            catch(Exception ex) {
                return JsonError(ex.Message);
            }
        }

        public ActionResult Projects(City? city, int page = 1)
        {
            var paging = new Page(page);
            var query = db.Projects.AsQueryable();
            if (city.HasValue)
            {
                query = query.Where(e => e.City == city.Value);
            }
            paging.RecordCount = query.Count();
            ViewBag.List = query.OrderBy(e => e.ID).Skip(paging.PageSize * (paging.PageIndex - 1)).Take(paging.PageSize).ToList();
            ViewBag.Page = paging;
            return View();
        }

        /// <summary>
        /// 上传总表数据
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects()
        {
            var file = UploadHelper.GetPostedFile(HttpContext);

            var list = new List<Project>();

            IWorkbook excel = null;
            var ext = Path.GetExtension(file.FileName);
            if (ext == ".xls")
            {
                excel = new HSSFWorkbook(file.InputStream);
            }
            else
            {
                excel = new XSSFWorkbook(file.InputStream);
            }
            //var excel = new HSSFWorkbook(file.InputStream);
            var sheet = excel.GetSheetAt(0);
            for (var i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (string.IsNullOrEmpty(row.Cells[0].StringCellValue))
                {
                    continue;
                }
                var cityNames = row.Cells[0].StringCellValue.Split(',');
                if (cityNames.Length < 3)
                {
                    continue;
                }

                var county = cityNames[2];

                City city = 0;

                if (Enum.TryParse<City>(cityNames[1], out city))
                {
                    list.Add(new Project
                    {
                        City = city,
                        ID = row.Cells[1].NumericCellValue.ToString(),
                        Name = row.Cells[2].StringCellValue,
                        County = county
                    });
                }

            }

            AddProjects(list);

            return RedirectToAction("Index", "Admin");
        }
    }
}