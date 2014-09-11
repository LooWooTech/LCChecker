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
                query = query.Skip(page.PageSize * (page.PageIndex - 1)).Take(page.PageSize);
            }

            return query.ToList();
        }


        public ActionResult Index(bool? result, int page = 1)
        {
            var paging = new Page(page);
            ViewBag.Projects = GetProjects(result, paging);
            ViewBag.Page = paging;
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
            Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();
            Dictionary<string, int> ship = new Dictionary<string, int>();
            DetectEngine Engine = new DetectEngine(filePath);
            string fault="";
            if (!Engine.CheckExcel(filePath, ref fault, ref Error, ref ship))
            {
                throw new ArgumentException("检索失败");
            }
            //检查完毕，更新Projects
            var projects = db.Projects.Where(x => x.City == CurrentUser.City).ToList();
            foreach (var item in projects)
            {
                if (item.Result==true)
                    continue;
                if (Error.ContainsKey(item.ID))
                {
                    item.Note = "";
                    foreach (var Message in Error[item.ID])
                    {
                        item.ID += Message + "；";
                    }
                    db.Entry(item).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }

            return RedirectToAction("Index?result=false");
        }
    }
}
