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
        private List<Project> GetProjects()
        {
            var query = db.Projects.Where(e => e.City == CurrentUser.City);

            return query.ToList();
        }


        public ActionResult Index(bool? result, int page = 1)
        {

            ViewBag.Projects = GetProjects();

            return View();
        }

        /*用户上传文件*/
        [HttpPost]
        public ActionResult FileUpload(FormCollection form)
        {
            if (Request.Files.Count == 0)
            {
                throw new ArgumentException("请选择文件上传");
            }

            HttpPostedFileBase file = Request.Files[0];
            string ext = Path.GetExtension(file.FileName);
            if (ext != ".xls" || ext != "xlsx")
            {
                throw new ArgumentException("你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格");
            }
            if (file.ContentLength == 0 || file.ContentLength > 20971520)
            {
                throw new ArgumentException("你上传的文件数据太大或者没有");
            }
            string filePath = null;
            Detect record = db.DETECT.FirstOrDefault(x => x.region == CurrentUser.name);
            if (record == null)
            {
                record = new Detect() { submit = 1, region = CurrentUser.name };
                db.DETECT.Add(record);
                db.SaveChanges();
            }
            else {
                record.submit++;
                db.Entry(record).State = EntityState.Modified;
                db.SaveChanges();
            }
            if (ext == ".xls")
            {
                filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + CurrentUser.name + "/" + record.submit), "NO" + record.submit + ".xls");
            }
            else
            {
                filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + CurrentUser.name + "/" + record.submit), "NO" + record.submit + ".xlsx");
            }
            string Catalogue = HttpContext.Server.MapPath("../Uploads/" + CurrentUser.name + "/" + record.submit);
            if (!Directory.Exists(Catalogue))
            {
                try
                {
                    Directory.CreateDirectory(Catalogue);
                }
                catch 
                {
                    throw new ArgumentException("创建目录失败");
                }
            }
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Create);
                fs.Close();
            }
            catch (Exception er)
            {
                throw new ArgumentException(er.Message);
            }
            file.SaveAs(filePath);
            SubRecord submitem = new SubRecord();
            submitem.Format = ext;
            submitem.regionId = record.Id;
            submitem.name = file.FileName;
            submitem.submits = record.submit;
            if (ModelState.IsValid)
            {
                db.SUBRECORD.Add(submitem);
                db.SaveChanges();
            }
            return RedirectToAction("");

            //return View();
        }
    }
}
