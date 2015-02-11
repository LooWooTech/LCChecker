using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Areas.Second.Controllers
{
    public class NewAreaController : SecondController
    {
        //
        // GET: /Second/NewArea/

        public ActionResult Index(NullableFilter result=NullableFilter.All,int page=1,string county=null)
        {
            var filter = new ProjectFileter
            {
                City=CurrentUser.City,
                Result=result,
                Page=new Page(page),
                County=county
            };
            ViewBag.List = ProjectHelper.GetCoordSeProjects(filter);
            ViewBag.Page = filter.Page;
            ViewBag.County = ProjectHelper.GetNewAreaCounty(CurrentUser.City);
            return View();
        }

        [HttpPost]
        public ActionResult Upload(UploadFileType type) {
            var file = HttpContext.GetPostedFile();
            if (file == null) {
                throw new ArgumentException("请选择需要上传的文件");
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
            return RedirectToAction("NewAreaUploadResult", new { type });
        }


        public ActionResult NewAreaUploadResult(UploadFileType type = UploadFileType.新增耕地坐标, int state = -1) {
            var query = db.Files.Where(e => e.City == CurrentUser.City && e.Type == type);
            if (state > -1) {
                query = query.Where(e => e.State == (UploadFileProceedState)state);
            }
            ViewBag.List = query.ToList();
            return View();
        }

        [HttpPost]
        public ActionResult AddException(string reason, string ID, int result=0, int page = 1, string county = null)
        {
            if (string.IsNullOrEmpty(reason)) {
                throw new ArgumentException("请输入添加例外理由！");
            }
            CoordNewAreaProject project = db.CoordNewAreaProjects.FirstOrDefault(e => e.ID.ToLower() == ID.ToLower());
            if (project == null) {
                throw new ArgumentException("未找到相关新增耕地坐标项目信息,请与管理员联系！");
            }
            project.Exception = true;
            project.Error = "例外理由：" + reason+";";
            db.SaveChanges();
            return RedirectToAction("Index", new { result,page,county});
        }


        public ActionResult CancelException(string ID, int result=0, int page = 1, string county = null)
        {
            CoordNewAreaProject project = db.CoordNewAreaProjects.FirstOrDefault(e => e.ID.ToLower() == ID.ToLower());
            if (project == null) {
                throw new ArgumentException("未找到相关新增耕地坐标项目信息，请与管理员联系！");
            }
            project.Exception = false;
            string[] Notes = project.Error.Split(';');
            string value = string.Empty;
            for (var i = 1; i < Notes.Length; i++) {
                value += Notes[i];
            }
            project.Error = value;
            db.SaveChanges();
            return RedirectToAction("Index", new { result,page,county});
        }

    }
}
