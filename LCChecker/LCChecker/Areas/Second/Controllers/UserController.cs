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


        public ActionResult Report(NullableFilter result=NullableFilter.All,int page=1,string PID=null,string country=null) {
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
                ID=PID,
                Country=country
            };
            ViewBag.Projects = SecondProjectHelper.GetProjects(filter);
            ViewBag.Page = filter.Page;
            ViewBag.Country = SecondProjectHelper.GetCountry(CurrentUser.City);
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
            foreach (var item in projects) {
                if (ids.Contains(item.ID)) {
                    if (errors.ContainsKey(item.ID))
                        item.Result = false;
                    else {
                        item.Result = true;
                    }
                }
            }
            db.SaveChanges();
            var list = db.SecondRecords.Where(e => e.City == CurrentUser.City && e.Type == type).ToList();
            if (type == SecondReportType.附表1)
            {
                var Seproject=engine.GetSeProject();
                foreach (var item in projects) {
                    if (!errors.ContainsKey(item.ID) && Seproject.ContainsKey(item.ID)) {
                        item.IsHasDoubt = Seproject[item.ID].IsHasDoubt;
                        item.IsApplyDelete = Seproject[item.ID].IsApplyDelete;
                        item.IsHasError = Seproject[item.ID].IsHasError;
                        item.IsPacket = Seproject[item.ID].IsPacket;
                        item.IsDescrease = Seproject[item.ID].IsDescrease;
                        item.IsRelieve = Seproject[item.ID].IsRelieve;
                        db.SaveChanges();
                    }
                }
            }
            if (type == SecondReportType.附表9) {
                var paddys = engine.GetPaddy();
                var drys = engine.GetDry();
                var farmland = db.FarmLands.ToList();
                foreach (var item in farmland) {
                    if (paddys.ContainsKey(item.ProjectID) && item.Type == LandType.Paddy) {
                        item.Area = paddys[item.ProjectID].Area;
                        item.Degree = paddys[item.ProjectID].Degree;
                        paddys.Remove(item.ProjectID);
                    }
                    if (drys.ContainsKey(item.ProjectID)&&item.Type==LandType.Dry) {
                        item.Area = drys[item.ProjectID].Area;
                        item.Degree = drys[item.ProjectID].Degree;
                        drys.Remove(item.ProjectID);
                    }
                }
                db.SaveChanges();
                foreach (var item in paddys.Keys) {
                    db.FarmLands.Add(new FarmLand() { ProjectID = item, Type = LandType.Paddy, Area = paddys[item].Area, Degree = paddys[item].Degree });
                }
                db.SaveChanges();
                foreach (var item in drys.Keys) {
                    db.FarmLands.Add(new FarmLand() { ProjectID = item, Type = LandType.Dry, Area = drys[item].Area, Degree = drys[item].Degree });
                }
                db.SaveChanges();
            }

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


        public ActionResult ReportResult(SecondReportType Type=0,RuleKind rule=RuleKind.All,bool? IsError=null,int page=1) {
            if (Type == 0) {
                throw new ArgumentException("参数错误，没有选择具体报部表格");
            }
            var query = db.SecondRecords.Where(e => e.City == CurrentUser.City && e.Type == Type);
            if (IsError != null) {
                query = query.Where(e => e.IsError == IsError.Value);
            }
            string str = string.Empty;
            switch (rule) {
                    
                case RuleKind.Basic:
                case RuleKind.Consistency:
                case RuleKind.Data:
                case RuleKind.Write:
                    str = rule.GetDescription();
                    query = query.Where(e => e.Note.Contains(str));
                    break;
                case RuleKind.Other:
                    str = RuleKind.Basic.GetDescription();
                    string one = RuleKind.Consistency.GetDescription();
                    string two = RuleKind.Data.GetDescription();
                    string three = RuleKind.Write.GetDescription();
                    query = query.Where(e => !(e.Note.Contains(str) || e.Note.Contains(one) || e.Note.Contains(two) || e.Note.Contains(three)));
                    break;
                case RuleKind.All:
                default: break;
            }
            var paging = new Page(page) { RecordCount = query.Count() };
            var list = query.OrderBy(e => e.ID).Skip(paging.PageSize * (paging.PageIndex - 1)).Take(paging.PageSize);
            ViewBag.Page = paging;
            ViewBag.Title = Type.GetDescription();
            return View(list);
        }

        

    }
}
