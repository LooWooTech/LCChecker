using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using LCChecker.Models;

namespace LCChecker.Controllers
{
    public partial class UserController
    {
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

        public ActionResult CoordProjectUploadResult(UploadFileType type = UploadFileType.项目坐标)
        {
            var query = db.Files.Where(e => e.City == CurrentUser.City && e.Type == type);

            query = query.Where(e => e.State == UploadFileProceedState.Error || e.State == UploadFileProceedState.UnProceed);

            ViewBag.List = query.ToList();
            return View();
        }
    }
}