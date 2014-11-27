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


        public ActionResult Report(NullableFilter result=NullableFilter.All,int page=1) {
            if (!db.SecondReports.Any(e => e.City == CurrentUser.City))
            {
                InitSecondReports();
            }
            ViewBag.List = db.SecondReports.Where(e => e.City == CurrentUser.City).ToList();
            var filter = new SecondProjectFilter
            {
                City = CurrentUser.City,
                Result = result,
                Page = new Page(page),
            };
            ViewBag.Projects = SecondProjectHelper.GetProjects(filter);
            ViewBag.Page = filter.Page;
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
            var file=db.Files.FirstOrDefault(e=>e.ID==ID);
            if (file == null)
            {
                throw new ArgumentException("参数错误");
            }
            else {
                file.State = UploadFileProceedState.Proceeded;
            }
            string Fault="";
            var filePath = UploadHelper.GetAbsolutePath(file.SavePath);
            List<SecondProject> projects = db.SecondProjects.Where(e => e.City == CurrentUser.City).ToList();
            ISeCheck engine = null;
            switch (type) { 
                case SecondReportType.附表1:
                    engine = new CheckOne(projects);
                    break;
                case SecondReportType.附表2:
                    engine = new CheckTwo(projects);
                    break;
                case SecondReportType.附表3:
                    engine = new CheckThree(projects);
                    break;
                case SecondReportType.附表4:
                    engine = new CheckFour(projects);
                    break;
                case SecondReportType.附表6:
                    engine = new CheckSix(projects);
                    break;
                case SecondReportType.附表7:
                    engine = new CheckSeven(projects);
                    break;
                case SecondReportType.附表8:
                    engine = new CheckEight(projects);
                    break;
                case SecondReportType.附表9:
                    engine = new CheckNine(projects);
                    break;
                default: file.State = UploadFileProceedState.Error;
                    file.ProcessMessage = "不支持当前业务类型" + (int)type;
                    db.SaveChanges();
                    throw new ArgumentException("不支持当前业务类型");
            }

            if (!engine.Check(filePath, ref Fault, type)) {
                file.State = UploadFileProceedState.Error;
                file.ProcessMessage = "检索失败：" + Fault;
                db.SaveChanges();
                throw new ArgumentException("检索表格失败"+Fault);
            }
            var errors = engine.GetError();
            var ids = engine.GetIDS();
            var list = db.SecondRecords.Where(e => e.City == CurrentUser.City && e.Type == type).ToList();
            SecondRecord.Clear(list);
            List<SecondRecord> records = new List<SecondRecord>();
            foreach (var item in errors.Keys) {
                Fault = "";
                int i = 1;
                foreach (var msg in errors[item]) {
                    Fault += string.Format("({0}){1}", i, msg);
                    i++;
                }
                records.Add(new SecondRecord()
                {
                    ProjectID=item,
                    Type=type,
                    City=CurrentUser.City,
                    IsError=true,
                    Note=Fault
                });
            }
            SecondRecord.AddRecords(records);
            var reports = db.SecondReports.FirstOrDefault(e => e.City == CurrentUser.City && e.Type == type);
            reports.Result = (errors.Count() == 0);
            if (errors.Count > 0) {
                file.State = UploadFileProceedState.Error;
            }
            db.SaveChanges();
            return RedirectToAction("ReportResult", new { Type = type });
           // return View();
        }


        public ActionResult ReportResult(SecondReportType Type=0,bool? IsError=null,int page=1) {
            if (Type == 0) {
                throw new ArgumentException("参数错误，没有选择具体报部表格");
            }
            var query = db.SecondRecords.Where(e => e.City == CurrentUser.City && e.Type == Type);
            if (IsError != null) {
                query = query.Where(e => e.IsError == IsError.Value);
            }
            var paging = new Page(page) { RecordCount = query.Count() };
            var list = query.OrderBy(e => e.ID).Skip(paging.PageSize * (paging.PageIndex - 1)).Take(paging.PageSize);
            ViewBag.Page = paging;
            ViewBag.Title = Type.GetDescription();
            return View(list);
        }

        

    }
}
