using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Areas.Second.Controllers
{
    public class UserController : SecondController
    {
        //
        // GET: /Second/User/

        public ActionResult Index()
        {
            return View();
        }


        public ActionResult Report() {
            if (!db.SecondReports.Any(e => e.City == CurrentUser.City))
            {
                InitSecondReports();
            }
            ViewBag.List = db.SecondReports.Where(e => e.City == CurrentUser.City).ToList();
            return View();
        }

        private void InitSecondReports() {
            foreach (var item in Enum.GetNames(typeof(SecondReportType))) {
                db.SecondReports.Add(new SecondReport
                {
                    City = CurrentUser.City,
                    Type = (SecondReportType)Enum.Parse(typeof(SecondReportType), item)
                });
                db.SaveChanges();
            }
        }


        [HttpPost]
        public ActionResult UploadReports(SecondReportType type) {
            var file = UploadHelper.GetPostedFile(HttpContext);
            var ext = Path.GetExtension(file.FileName);
            if (ext != ".xls" && ext != ".xlsx") {
                throw new ArgumentException("你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格");
            }
            var filePath = UploadHelper.Upload(file);
            var fileID = UploadHelper.AddFileEntity(new UploadFile
            {
                City = CurrentUser.City,
                CreateTime = DateTime.Now,
                FileName = file.FileName,
                SavePath = filePath,
                Type = (UploadFileType)((int)type + 20)
            });
            var record = db.SecondReports.FirstOrDefault(e => e.City == CurrentUser.City && e.Type == type);
            if (record != null) {
                record.Result = false;
                db.SaveChanges();
            }
            return RedirectToAction("CheckReport", new { ID = fileID, type = type });
        }


        public ActionResult CheckReport(int ID, SecondReportType type) {
            return View();
        }

    }
}
