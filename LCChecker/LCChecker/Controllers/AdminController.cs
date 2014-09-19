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

        public ActionResult Projects(City? city, NullableFilter result = NullableFilter.All, int page = 1)
        {
            var filter = new ProjectFileter
            {
                City = city,
                Result = result,
                Page = new Page(page)
            };
            ViewBag.List = ProjectHelper.GetProjects(filter);
            ViewBag.Page = filter.Page;
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

            ViewBag.CoordSummary = db.CoordProjects.Where(e => e.Visible == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
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

            int rowIndex = 0, cellIndex = 0;
            var title = sheet.GetRow(rowIndex).GetCell(cellIndex).GetValue();
            if (title != "附表3")
            {
                throw new ArgumentException("上传的附表3：重点项目复核确认总表格式不正确，请参照样表。");
            }
            rowIndex++;
            for (var i = rowIndex+5; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                if (string.IsNullOrEmpty(row.Cells[cellIndex].ToString()))
                {
                    continue;
                }
                var id = row.Cells[cellIndex + 3].GetValue().Trim();
                //if (!id.VerificationID())
                //{
                //    continue;
                //}

                City city = 0;

                if (Enum.TryParse<City>(row.Cells[cellIndex + 1].GetValue(), out city))
                {
                    list.Add(new Project
                    {
                        City = city,
                        County = row.Cells[cellIndex + 2].GetValue(),
                        ID = id,
                        Name = row.Cells[cellIndex + 4].GetValue(),
                        IsHasError = row.Cells[cellIndex + 5].GetValue() == "否",
                        IsApplyDelete = row.Cells[cellIndex + 6].GetValue() == "是",
                        IsShouldModify = row.Cells[cellIndex + 7].GetValue() == "是",
                        IsDecrease = row.Cells[cellIndex + 9].GetValue() == "是",

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
                Visible = true,
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

            var visible = false;

            var firstCell = sheet.GetRow(0).GetCell(0);
            if (firstCell.GetValue() == "市")
            {
                visible = true;
            }

            var rowIndex = visible ? 1 : 7;
            var cellIndex = visible ? 0 : 1;
            for (; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (string.IsNullOrEmpty(row.Cells[0].ToString()))
                {
                    continue;
                }

                if (!visible && row.Cells[6].StringCellValue == "是")
                {
                    continue;
                }

                City city = 0;

                if (Enum.TryParse<City>(row.Cells[cellIndex].GetValue(), out city))
                {
                    list.Add(new CoordProject
                    {
                        City = city,
                        County = row.Cells[cellIndex + 1].GetValue(),
                        ID = row.Cells[cellIndex + 2].GetValue(),
                        Name = row.Cells[cellIndex + 3].GetValue(),
                        Note = visible ? row.Cells[cellIndex + 4].GetValue() : null,
                        Visible = visible
                    });
                }
            }
            ProjectHelper.AddCoordProjects(list);
            return RedirectToAction("CoordProjects");
        }
    }
}