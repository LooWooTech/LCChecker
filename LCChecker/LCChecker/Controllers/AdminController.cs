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
        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (!CurrentUser.Flag)
            {
                throw new HttpException(401, "你没有权限查看此页面");
            }
            base.OnActionExecuting(filterContext);
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
            catch (Exception ex)
            {
                return JsonError(ex.Message);
            }
        }

        public ActionResult Projects(NullableFilter result = NullableFilter.All, int page = 1)
        {
            var filter = new ProjectFileter
            {
                City = CurrentUser.City,
                Result = result,
                Page = new Page(page)
            };
            ViewBag.List = ProjectHelper.GetProjects(filter);
            ViewBag.Page = filter.Page;
            return View();
        }

        public ActionResult ReportSummary()
        {
            ViewBag.Summary = db.Reports.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            return View();
        }

        [HttpGet]
        public ActionResult Statistics()
        {
            ViewBag.ProjectSummary = db.Projects.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });

            ViewBag.ReportSummary = db.Reports.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });

            ViewBag.CoordSummary = db.CoordProjects.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            return View();
        }

        /// <summary>
        /// 上传总表数据（导入表3）
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects()
        {
            var file = UploadHelper.GetPostedFile(HttpContext);

            var list = new List<Project>();

            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(0);

            int startRow = 0, startCell = 0;
            if (!sheet.FindHeader(ref startRow, ref startCell, ReportType.附表3))
            {
                throw new ArgumentException("未找到重点项目复核确认总表的表头");
            }
            startRow++;
            for (var i = startRow; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (string.IsNullOrEmpty(row.Cells[0].ToString()))
                {
                    continue;
                }
                var value = row.GetCell(startCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                if (!value.VerificationID())
                    continue;

                City city = 0;

                if (Enum.TryParse<City>(row.Cells[1].ToString(), out city))
                {
                    list.Add(new Project
                    {
                        City = city,
                        County = row.Cells[2].StringCellValue,
                        ID = row.Cells[3].NumericCellValue.ToString(),
                        Name = row.Cells[4].StringCellValue,
                        IsHasError = row.Cells[5].StringCellValue == "否",
                        IsApplyDelete = row.Cells[6].StringCellValue == "是",
                        IsShouldModify = row.Cells[7].StringCellValue == "是",
                        IsDecrease = row.Cells[9].StringCellValue == "是",

                    });
                }

            }

            ProjectHelper.AddProjects(list);

            return RedirectToAction("Projects", "Admin");
        }

        public ActionResult CoordProjects(NullableFilter result = NullableFilter.All, int page = 1)
        {
            var filter = new ProjectFileter
            {
                City = CurrentUser.City,
                Result = result,
                Page = new Page(page)
            };
            ViewBag.List = ProjectHelper.GetCoordProjects(filter);
            ViewBag.Page = filter.Page;
            return View();
        }

        [HttpPost]
        public ActionResult UploadCoords()
        {
            var file = UploadHelper.GetPostedFile(HttpContext);


            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(0);

            var list = new List<CoordProject>();
            for (var i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (string.IsNullOrEmpty(row.Cells[0].ToString()))
                {
                    continue;
                }
                City city = 0;

                if (Enum.TryParse<City>(row.Cells[0].ToString(), out city))
                {
                    list.Add(new CoordProject
                    {
                        City = city,
                        County = row.Cells[1].StringCellValue,
                        ID = row.Cells[2].StringCellValue,
                        Name = row.Cells[3].StringCellValue,
                        Note = row.Cells[4].StringCellValue
                    });
                }
            }
            ProjectHelper.AddCoordProjects(list);
            return RedirectToAction("CoordProjects");
        }
    }
}