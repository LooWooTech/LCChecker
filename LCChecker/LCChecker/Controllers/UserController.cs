using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    [UserAuthorize]
    public class UserController : ControllerBase
    {
        private List<Project> GetProjects(bool? result = null, Page page = null)
        {
            var query = db.Projects.Where(e => e.City == CurrentUser.City);

            if (result.HasValue)
            {
                query = query.Where(e => e.Result == result.Value);
            }

            if (page != null)
            {
                page.RecordCount = query.Count();
                query = query.OrderBy(e => e.ID).Skip(page.PageSize * (page.PageIndex - 1)).Take(page.PageSize);
            }

            return query.ToList();
        }


        public ActionResult Index(bool? result, int page = 1)
        {
            var summary = new Summary
            {
                City = CurrentUser.City,
                TotalCount = db.Projects.Count(e => e.City == CurrentUser.City),
                SuccessCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == true),
                ErrorCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == false),

            };
            //全部结束，进入第二阶段
            if (summary.TotalCount == summary.SuccessCount)
            {

                return View("Index2");
            }

            var paging = new Page(page);
            ViewBag.Projects = GetProjects(result, paging);
            ViewBag.Page = paging;
            ViewBag.Summary = summary;
            return View();
        }

        /// <summary>
        /// 下载未完成和错误的Project模板
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadProjects()
        {
            var list = db.Projects.Where(e => e.Result != null).ToList();
            return View();
        }

        /// <summary>
        /// 上传一部分项目，验证并更新到Project
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects(FormCollection form)
        {
            var file = UploadHelper.GetPostedFile(HttpContext);

            var filePath = UploadHelper.UploadExcel(file);

            var uploadFile = new UploadFile
            {
                City = CurrentUser.City,
                CreateTime = DateTime.Now,
                FileName = file.FileName,
                SavePath = filePath
            };

            db.Files.Add(uploadFile);

            db.SaveChanges();

            //上传成功后跳转到check页面进行检查，参数是File的ID
            return RedirectToAction("Check", new { id = uploadFile.ID });
        }

        public ActionResult Check(int id)
        {
            var file = db.Files.FirstOrDefault(e => e.ID == id);
            if (file == null)
            {
                throw new ArgumentException("参数错误");
            }

            var filePath = file.SavePath;
            //读取文件进行检查

            //检查完毕，更新Projects

            return RedirectToAction("Index?result=false");
        }
    }
}
