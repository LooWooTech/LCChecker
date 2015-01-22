using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using NPOI.SS.UserModel;
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
            if (!db.SecondReports.Any(e => e.City == CurrentUser.City&&e.IsPlan==false))
            {
                InitSecondReports(false);
            }
            ViewBag.List = db.SecondReports.Where(e => e.City == CurrentUser.City&&e.IsPlan==false).ToList();
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
            ViewBag.Country = SecondProjectHelper.GetCountry(CurrentUser.City,false);
            return View();
        }

        private void InitSecondReports(bool isPlan) {
            foreach (SecondReportType item in Enum.GetValues(typeof(SecondReportType))) {
                if (isPlan&&item == SecondReportType.附表6) {
                    return;
                }
                db.SecondReports.Add(new SecondReport
                {
                    City = CurrentUser.City,
                    Type = item,
                    IsPlan = isPlan
                });
                db.SaveChanges();
            }


            //foreach (var item in Enum.GetNames(typeof(SecondReportType))) {
                
            //    db.SecondReports.Add(new SecondReport
            //    {
            //        City = CurrentUser.City,
            //        Type = (SecondReportType)Enum.Parse(typeof(SecondReportType), item),
            //        IsPlan=isPlan
            //    });
            //    db.SaveChanges();
            //}
        }

        public ActionResult planReport(NullableFilter result=NullableFilter.All,string PID=null,int page=1,string country=null) {
            if(!db.SecondReports.Any(e=>e.City==CurrentUser.City&&e.IsPlan==true)){
                InitSecondReports(true);
            }
            ViewBag.List = db.SecondReports.Where(e => e.City == CurrentUser.City && e.IsPlan == true).ToList();
            var filter = new SecondProjectFilter
            {
                City = CurrentUser.City,
                Result = result,
                Page = new Page(page),
                ID=PID,
                Country = country
            };
            ViewBag.pProjects=SecondProjectHelper.GetPlanProjects(filter);
            ViewBag.Page=filter.Page;
            ViewBag.county = SecondProjectHelper.GetCountry(CurrentUser.City, true);
            return View();
        }


        [HttpPost]
        public ActionResult UploadReports(SecondReportType type,bool IsPlan=false) {
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
                Type = (UploadFileType)((int)type + 20),
                IsPlan=IsPlan
            });
            if (IsPlan)
            {
                return RedirectToAction("CheckPlanProject", new { ID = fileID, type = type });
            }
            return RedirectToAction("CheckReport", new { ID = fileID, type = type});
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
           
            ISeCheck engine = null;
            #region
            List<SecondProject> projects = db.SecondProjects.Where(e => e.City == CurrentUser.City).ToList();
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

            #region
            if (!engine.Check(filePath, ref Fault, type,false)) {
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
            var list = db.SecondRecords.Where(e => e.City == CurrentUser.City && e.Type == type&&e.IsPlan==false).ToList();
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
            #endregion
            #region
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
            #endregion

            #region
            SecondRecord.Clear(list);
            List<SecondRecord> records = new List<SecondRecord>();
            int Count = engine.GetNumber();
            if (Count != 0) {
                records.Add(new SecondRecord {
                    ProjectID = "表格",
                    Type = type,
                    City = CurrentUser.City,
                    IsError = true,
                    Note = "本次检查有"+Count+"个项目未检查,请核对项目是否包括在验收项目中",
                    IsPlan = false
                });
            }
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
                    Note=Fault,
                    IsPlan=false
                });
            }
            #endregion
            SecondRecord.AddRecords(records);
            var reports = db.SecondReports.FirstOrDefault(e => e.City == CurrentUser.City && e.Type == type&&e.IsPlan==false);
            reports.Result = ((errors.Count()+Count) == 0);
            if (errors.Count > 0) {
                file.State = UploadFileProceedState.Error;
            }
            db.SaveChanges();
            #endregion
            return RedirectToAction("ReportResult", new { Type = type,IsPlan=false});
           // return View();
        }

        public ActionResult CheckPlanProject(int ID, SecondReportType Type)
        {
            var file = db.Files.FirstOrDefault(e => e.ID == ID);
            if (file == null)
            {
                throw new ArgumentException("参数错误");
            }
            else
            {
                file.State = UploadFileProceedState.Proceeded;
            }
            string Fault = "";
            var filePath = UploadHelper.GetAbsolutePath(file.SavePath);
            //初始化未验收项目检查引擎
            #region
            ISePlanCheck engine = null;
            List<pProject> projects = db.pProjects.Where(e => e.City == CurrentUser.City).ToList();
            switch (Type)
            {
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
                default: file.State = UploadFileProceedState.Error;
                    file.ProcessMessage = "不支持当前业务类型" + (int)Type;
                    db.SaveChanges();
                    throw new ArgumentException("不支持当前业务类型");
            }
            #endregion
            //开始进行检查
            if (!engine.Check(filePath, ref Fault, Type, true))
            {
                file.State = UploadFileProceedState.Error;
                file.ProcessMessage = "检索失败：" + Fault;
                db.SaveChanges();
                throw new ArgumentException("检索表格失败" + Fault);
            }
            //获取检查结果
            var errors = engine.GetError();

            //获取数据库中存在对应报部表格检查结果
            var list = db.SecondRecords.Where(e => e.City == CurrentUser.City && e.Type == Type && e.IsPlan == true).ToList();
            //附表1需要对数据库中的字段进行更新
            if (Type == SecondReportType.附表1)
            {
                var SePlanProject = engine.GetPlanData();
                foreach (var item in projects) {
                    item.IsRight = false;
                    item.IsApplyDelete = false;
                    item.IsHasError = false;
                    item.IsHad = false;
                }
                db.SaveChanges();
                foreach (var item in SePlanProject) {
                    var pproject = db.pProjects.Where(e => e.Key.Trim().ToUpper() == item.Key.Trim().ToUpper() && e.Name.Trim().ToUpper() == item.Name.Trim().ToUpper() && e.County.Trim().ToUpper() == item.County.Trim().ToUpper()&&e.IsHad==false).FirstOrDefault();
                    if (pproject != null) {
                        pproject.IsHad = true;
                        pproject.IsRight = item.IsRight;
                        pproject.IsHasError = item.IsHasError;
                        pproject.IsApplyDelete = item.IsApplyDelete;
                        db.SaveChanges();
                    }
                }
                db.SaveChanges();
            }
            SecondReport reports = db.SecondReports.FirstOrDefault(e => e.City == CurrentUser.City && e.Type == Type && e.IsPlan);
            reports.Note = engine.GetNumber().ToString();
            db.SaveChanges();
            List<SecondRecord> records = new List<SecondRecord>();
            if (Type == SecondReportType.附表4)
            {
                if (!SecondProjectHelper.Check(CurrentUser.City))
                {
                    records.Add(new SecondRecord
                    {
                        ProjectID = "全部",
                        County = "全部",
                        Name = "全部",
                        Type = Type,
                        City = CurrentUser.City,
                        IsError = true,
                        Note = "附表1检查项目等于附表2、附表3以及附表4中的项目之和",
                        IsPlan = true
                    });
                }
            }
            if (errors.Count > 0||records.Count>0)
            {
                reports.Result = false;
            }
            else
            {
                reports.Result = true;
            }
            
            db.SaveChanges();
           
            //更新数据库中相关报部表格中记录  首先需要对之前数据库中存在清空
            SecondRecord.Clear(list);
            //获取本次报部表格检查结果，并且声称List
            
           
            foreach (var item in errors.Keys)
            {
                Fault = "";
                int i = 1;
                foreach (var msg in errors[item])
                {
                    Fault += string.Format("({0}){1}", i, msg);
                    i++;
                }
                var values = item.Split('-');
                records.Add(new SecondRecord()
                {
                    ProjectID = values[2].Trim(),
                    County = values[1].Trim(),
                    Name = values[0].Trim(),
                    Type = Type,
                    City = CurrentUser.City,
                    IsError = true,
                    Note = Fault,
                    IsPlan = true
                });
            }
            //添加本次检查结果
            SecondRecord.AddRecords(records);


            
           
            if (errors.Count > 0)
            {
                file.State = UploadFileProceedState.Error;
            }

            db.SaveChanges();
            
            return RedirectToAction("ReportResult", new { Type = Type, IsPlan = true });
        }



        public ActionResult ReportResult(bool IsPlan,SecondReportType Type=0,RuleKind rule=RuleKind.All,string County=null,int page=1) {
            if (Type == 0) {
                throw new ArgumentException("参数错误，没有选择具体报部表格");
            }
            var query = db.SecondRecords.Where(e => e.City == CurrentUser.City && e.Type == Type&&e.IsPlan==IsPlan);
            List<string> Countys = query.GroupBy(e => e.County.ToUpper()).Select(g => g.Key).ToList();
            ViewBag.Countys = Countys;
            if (!string.IsNullOrEmpty(County)) {
                query = query.Where(e => e.County.ToUpper() == County.ToUpper());
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
            ViewBag.IsPlan = IsPlan;
            return View(list);
        }

        public ActionResult DownLoadReason(bool IsPlan) {
            List<string[]> Data = new List<string[]>();
            List<CoordProjectBase> projects = new List<CoordProjectBase>();
            string HeadName = string.Empty;
            if (IsPlan)
            {
                HeadName = CurrentUser.City.ToString()+"新增耕地坐标存疑项目清单";
                projects = db.CoordNewAreaProjects.Where(e => e.Result == false && e.Exception == false && e.City == CurrentUser.City).Select(e => new CoordProjectBase
                {
                    ID=e.ID,
                    City= e.City,
                    County= e.County,
                    Name= e.Name,
                    Note= e.Note
                }).ToList();
            }
            else {
                HeadName = CurrentUser.City.ToString() + "坐标存疑项目清单";
                projects = db.CoordProjects.Where(e => e.Result == false && e.Visible == true && e.Exception == false && e.City == CurrentUser.City).Select(e => new CoordProjectBase
                {
                    ID= e.ID,
                    City= e.City,
                    County= e.County,
                    Name= e.Name,
                    Note= e.Note
                }).ToList();
               
            }
            foreach (var item in projects) {
                string[] values = new string[5] { 
                    item.ID,
                    item.City.ToString(),
                    item.County,
                    item.Name,
                    item.Note
                };
                Data.Add(values);
            }
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/Second/Mistakes.xls");
            IWorkbook workbook = null;
            try
            {
                using (var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read)) {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch (Exception er){
                throw new ArgumentException(er.ToString());
            }
            int SerialNumber = 1;
            int StartRow = 3;
            ISheet sheet = workbook.GetSheetAt(0);
            IRow TemplateRow = sheet.GetRow(StartRow);
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            cell.SetCellValue(HeadName);
            foreach (var array in Data) {
                row = sheet.GetRow(StartRow);
                if (row == null) {
                    row = sheet.CreateRow(StartRow);
                    if (TemplateRow.RowStyle != null) {
                        row.RowStyle = TemplateRow.RowStyle;
                    }
                   
                }
                cell = row.GetCell(0);
                if (cell == null) {
                    cell = row.CreateCell(0,TemplateRow.GetCell(0).CellType);
                    cell.CellStyle = TemplateRow.GetCell(0).CellStyle;
                }
                cell.SetCellValue(SerialNumber);
                SerialNumber++;
                StartRow++;
                int TempNumber = 1;
                foreach (var item in array) {
                    cell = row.GetCell(TempNumber);
                    if (cell == null) {
                        cell = row.CreateCell(TempNumber, TemplateRow.GetCell(TempNumber).CellType);
                        cell.CellStyle = TemplateRow.GetCell(TempNumber).CellStyle;
                    }
                    cell.SetCellValue(item);
                    TempNumber++;
                }

            }
            using (var ms = new MemoryStream()) {
                workbook.Write(ms);
                return File(ms.ToArray(), "application/ms-excel", HeadName + ".xls");
            }
            //return File();
        }

        

    }
}
