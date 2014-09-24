using LCChecker.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    public partial class UserController
    {
        public ActionResult Reports()
        {
            //var projects = new Summary
            //{
            //    City = CurrentUser.City,
            //    TotalCount = db.Projects.Count(e => e.City == CurrentUser.City),
            //    SuccessCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == true),
            //    ErrorCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == false)
            //};
            //ViewBag.Flag = (projects.TotalCount == projects.SuccessCount);
            ViewBag.Flag = true;
            if (!db.Reports.Any(e => e.City == CurrentUser.City))
            {
                InitReports();
            }
            ViewBag.List = db.Reports.Where(e => e.City == CurrentUser.City).ToList();
            return View();
        }

        private void InitReports()
        {
            int i = 0;
            foreach (var item in Enum.GetNames(typeof(ReportType)))
            {
                db.Reports.Add(new Report
                {
                    ID = (int)CurrentUser.City + "_" + i.ToString(),
                    City = CurrentUser.City,
                    Type = (ReportType)Enum.Parse(typeof(ReportType), item)
                });
                db.SaveChanges();
                i++;
            }
        }

        public ActionResult UploadReport(ReportType type)
        {
            var file = UploadHelper.GetPostedFile(HttpContext);

            var ext = Path.GetExtension(file.FileName);
            if (ext != ".xls" && ext != ".xlsx")
            {
                throw new ArgumentException("你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格");
            }
            var filePath = UploadHelper.Upload(file);
            var fileId = UploadHelper.AddFileEntity(new UploadFile
            {
                City = CurrentUser.City,
                CreateTime = DateTime.Now,
                FileName = file.FileName,
                SavePath = filePath,
                Type = (UploadFileType)(int)type
            });

            var record = db.Reports.FirstOrDefault(e => e.City == CurrentUser.City && e.Type == type);
            if (record != null)
            {
                record.Result = false;
                db.SaveChanges();
            }
            return RedirectToAction("CheckReport", new { id = fileId, type = (int)type });
        }

        public ActionResult CheckReport(int id, ReportType type)
        {
            var file = db.Files.FirstOrDefault(e => e.ID == id);
            if (file == null)
            {
                throw new ArgumentException("参数错误");
            }
            else
            {
                file.State = UploadFileProceedState.Proceeded;
            }

            var filePath = UploadHelper.GetAbsolutePath(file.SavePath);
            string MastPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/", CurrentUser.City.ToString() + ".xls");
            List<Project> projects = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            bool Flag = false;
            ICheck engine = null;
            switch (type)
            {
                case ReportType.附表4:
                    engine = new CheckReport4(projects);
                    Flag = true;
                    break;
                case ReportType.附表5:
                    engine = new CheckReport5(projects);
                    Flag = true;
                    break;
                case ReportType.附表7:
                    engine = new CheckReport7(MastPath,projects);
                    break;
                case ReportType.附表8:
                    engine = new CheckReport8(MastPath,projects);
                    Flag = true;
                    break;
                case ReportType.附表9:
                    engine = new CheckReport9(MastPath,projects);
                    break;
                default:
                    file.State = UploadFileProceedState.Error;
                    file.ProcessMessage = "不支持当前业务类型：" + (int)type;
                    db.SaveChanges();
                    throw new ArgumentException("不支持当前业务类型");
            }

            string fault = "";
            if (!engine.Check(filePath, ref fault, type, projects, Flag))
            {
                file.State = UploadFileProceedState.Error;
                file.ProcessMessage = "检索失败：" + fault;
                db.SaveChanges();
                throw new ArgumentException("检索表格失败" + fault);
            }

            var errors = engine.GetError();
            var warning = engine.GetWarning();
            var ids = engine.GetIDS();
            var list = db.Records.Where(x => x.City == CurrentUser.City && x.Type == type).ToList();
            DeleteRecords(list);
            List<Record> Records = new List<Record>();
            foreach (var item in errors.Keys)
            {
                fault = "";
                fault = errors[item];
                //int i = 1;
                //foreach (var msg in errors[item])
                //{
                //    fault += string.Format("（{0}）{1}；", i, msg);
                //    i++;
                //}
                Records.Add(new Record()
                {
                    ProjectID = item,
                    Type = type,
                    City = CurrentUser.City,
                    IsError = true,
                    Note = fault
                });
            }
            foreach (var item in warning.Keys)
            {
                Records.Add(new Record()
                {
                    ProjectID = item,
                    Type = type,
                    City = CurrentUser.City,
                    IsError = false,
                    Note=warning[item]
                    //Note = string.Format("(1){0}", warning[item])
                });
            }
            AddRecords(Records);


            var record = db.Reports.FirstOrDefault(e => e.City == CurrentUser.City && e.Type == type);
            //当提交的表格没有错误并且没有提示的时候 设置标志位为true 有可能没有错误 但是有提示，那个时候可以让用户上传
            record.Result = ((errors.Count == 0) && (warning.Count == 0));
            if (errors.Count > 0)
            {
                file.State = UploadFileProceedState.Error;
            }
            db.SaveChanges();
            ViewBag.Warning = warning;

            return RedirectToAction("ReportResult", new { type });
            //return View(errors);
        }

        public void DeleteRecords(List<Record> list)
        {
            using (var db = new LCDbContext())
            {
                foreach (var item in list)
                {
                    db.Records.Attach(item);
                    db.Records.Remove(item);
                }
                db.SaveChanges();
            }
        }

        public static void AddRecords(List<Record> list)
        {
            using (var db = new LCDbContext())
            {
                foreach (var item in list)
                {
                    db.Records.Add(item);
                }
                try
                {
                    db.SaveChanges();
                }
                catch (DbEntityValidationException dbEx)
                { 
                
                }
                
            }


        }

        public ActionResult DownloadReport(ReportType type)
        {
            var workbook = XslHelper.GetWorkbook("Templates/" + type.ToString() + ".xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + type.GetDescription());

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            switch (type)
            {
                case ReportType.附表4:
                case ReportType.附表8:
                    GetReportExcel4(sheet, list, 6);
                    break;
                case ReportType.附表5:
                case ReportType.附表7:
                    GetReportExcel5(sheet, list, 5);
                    break;
                case ReportType.附表9:
                    GetReportExcel9(sheet, list);
                    break;
            }

            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                return File(ms.ToArray(), "application/ms-excel", type.ToString() + ".xls");
            }
        }

        private void GetReportExcel4(ISheet sheet, List<Project> list, int rowIndex)
        {
            var rowNumber = 1;
            sheet.InsertRow(rowIndex, list.Count - 1);
            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);

                rowIndex++;
                rowNumber++;
            }
        }

        private void GetReportExcel5(ISheet sheet, List<Project> list, int rowIndex)
        {
            var rowNumber = 1;
            sheet.InsertRow(rowIndex, list.Count - 1);
            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);
                row.Cells[3].SetCellValue(item.ID);
                row.Cells[4].SetCellValue(item.Name);
                row.Cells[8].SetCellValue("是");
                rowIndex++;
                rowNumber++;
            }
        }

        private void GetReportExcel9(ISheet sheet, List<Project> list)
        {
            var rowIndex = 6;
            var rowNumber = 1;

            var dkNames = new[] { "水田", "水浇地", "旱地" };
            sheet.InsertRow(rowIndex, list.Count * dkNames.Length);

            foreach (var item in list)
            {
                for (int j = 0; j < 6; j++)
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress((rowIndex - 3), (rowIndex - 1), j, j));
                }

                var groupRow = sheet.GetRow(rowIndex - 3);
                groupRow.Cells[0].SetCellValue(rowNumber.ToString());
                groupRow.Cells[1].SetCellValue(item.City.ToString());
                groupRow.Cells[2].SetCellValue(item.County);
                groupRow.Cells[3].SetCellValue(item.ID);
                groupRow.Cells[4].SetCellValue(item.Name);

                foreach (var name in dkNames)
                {
                    var row = sheet.GetRow(rowIndex);
                    row.Cells[6].SetCellValue(dkNames[rowIndex - 5]);
                    for (int i = 7; i < 22; i++)
                    {
                        row.Cells[i].SetCellValue("（亩）");
                    }
                    row.Cells[22].SetCellValue("是");
                    rowIndex++;
                }

                rowNumber++;
            }
        }

        public ActionResult ReportResult(ReportType type = 0, bool? isError = null, int page = 1)
        {
            if (type == 0)
            {
                throw new ArgumentException("参数错误，没有选择具体报部表");
            }

            var query = db.Records.Where(x => x.City == CurrentUser.City && x.Type == type);
            if (isError != null)
            {
                query = query.Where(e => e.IsError == isError.Value);
            }

            var paging = new Page(page) { RecordCount = query.Count() };
            var list = query.OrderBy(e => e.ID).Skip(paging.PageSize * (paging.PageIndex - 1)).Take(paging.PageSize);

            ViewBag.Page = paging;
            ViewBag.Title = type.GetDescription();

            return View(list);
        }
    }
}
