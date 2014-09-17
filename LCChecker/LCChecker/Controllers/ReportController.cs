using LCChecker.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
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

            var filePath = UploadHelper.UploadExcel(file);
            var fileId = UploadHelper.AddFileEntity(new UploadFile
            {
                City = CurrentUser.City,
                CreateTime = DateTime.Now,
                FileName = file.FileName,
                SavePath = filePath
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

            var filePath = UploadHelper.GetAbsolutePath(file.SavePath);
            string MastPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/", CurrentUser.City.ToString() + ".xls");
            ICheck engine = null;
            switch (type)
            {
                case ReportType.附表四:
                    engine = new CheckReport4(MastPath);
                    break;
                case ReportType.附表五:
                    engine = new CheckReport5(MastPath);
                    break;
                case ReportType.附表七:
                    engine = new CheckReport7(MastPath);
                    break;
                case ReportType.附表八:
                    engine = new CheckReport8(MastPath);
                    break;
                case ReportType.附表九:
                    engine = new CheckReport9(MastPath);
                    break;
                default:
                    throw new ArgumentException("不支持当前业务类型");
            }

            string fault = "";
            if (!engine.Check(filePath, ref fault))
            {
                throw new ArgumentException("检索表格失败" + fault);
            }

            var errors = engine.GetError();

            var record = db.Reports.FirstOrDefault(e => e.City == CurrentUser.City && e.Type == type);
            record.Result = errors.Count == 0;
            db.SaveChanges();

            return View(errors);
        }

        public ActionResult DownloadReport(ReportType type)
        {
            var workbook = XslHelper.GetWorkbook("Templates/" + (int)type + ".xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + type.GetDescription());

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            switch (type)
            {
                case ReportType.附表四:
                case ReportType.附表八:
                    GetReportExcel4(sheet, list, 6);
                    break;
                case ReportType.附表五:
                case ReportType.附表七:
                    GetReportExcel5(sheet, list, 5);
                    break;
                case ReportType.附表九:
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
    }
}
