using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using LCChecker.Models;
using LCChecker.Areas.Second.Models;

namespace LCChecker.Controllers
{
    public partial class UserController
    {
        public ActionResult CoordProjects(NullableFilter result = NullableFilter.All, int page = 1,string county=null)
        {
            var filter = new ProjectFileter
            {
                City = CurrentUser.City,
                Result = result,
                Visible=true,
                Page = new Page(page),
                County=county
            };
            ViewBag.List = ProjectHelper.GetCoordProjects(filter);
            ViewBag.Page = filter.Page;
            ViewBag.County = ProjectHelper.GetCoordCounty(CurrentUser.City);
            return View();
        }


        public ActionResult UploadCoordProjects(UploadFileType type)
        {
            var file = HttpContext.GetPostedFile();
            if (file == null)
            {
                throw new ArgumentNullException("请选择需要上传的文件");
            }

            var fileExt = System.IO.Path.GetExtension(file.FileName);
            if (fileExt != ".zip")
            {
                throw new ArgumentException("请上传zip格式的文件");
            }

            var savePath = file.Upload();

            UploadHelper.AddFileEntity(new UploadFile
            {
                City = CurrentUser.City,
                FileName = file.FileName,
                SavePath = savePath,
                Type = type
            });

            return RedirectToAction("CoordProjectUploadResult", new { type });

        }

        public ActionResult CoordProjectUploadResult(UploadFileType type = UploadFileType.项目坐标, int state = -1)
        {
            var query = db.Files.Where(e => e.City == CurrentUser.City && e.Type == type);
            if (state > -1)
            {
                query = query.Where(e => e.State == (UploadFileProceedState)state);
            }
            ViewBag.List = query.ToList();
            return View();
        }

        [HttpPost]
        public ActionResult AddException(string reason, string ID) {
            if (string.IsNullOrEmpty(reason)) {
                throw new ArgumentException("请输入添加例外理由！");
            }
            CoordProject project = db.CoordProjects.FirstOrDefault(e => e.ID == ID);
            if (project == null) {
                throw new ArgumentException("未找到相关坐标点项目信息，添加例外失败！请与管理员联系！");
            }
            project.Exception = true;
            project.Note = "例外理由：" + reason;
            db.SaveChanges();
            return RedirectToAction("CoordProjects");
        }
    }
}