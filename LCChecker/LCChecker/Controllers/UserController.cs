using LCChecker.Models;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
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
        private List<Project> GetProjects(bool? result = null, Page page = null)
        {
            var query = db.Projects.Where(e => e.City == CurrentUser.City);

            if (result.HasValue)
            {
                query = query.Where(e => e.Result == result.Value);
            }
            else
            {
                query = query.Where(e => e.Result == null);
            }
            if (page != null)
            {
                page.RecordCount = query.Count();
                query = query.OrderBy(e => e.ID).Skip(page.PageSize * (page.PageIndex - 1)).Take(page.PageSize);
            }

            return query.ToList();
        }

        public ActionResult Index(bool? result, int page = 1)
        {
            var summary = new Summary
            {
                City = CurrentUser.City,
                TotalCount = db.Projects.Count(e => e.City == CurrentUser.City),
                SuccessCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == true),
                ErrorCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == false),

            };
            //全部结束，进入第二阶段
            if (summary.TotalCount > 0 && summary.TotalCount == summary.SuccessCount)
            {

                return View("Index2");
            }

            var paging = new Page(page);
            ViewBag.Projects = GetProjects(result, paging);
            ViewBag.Page = paging;
            ViewBag.Summary = summary;
            return View();
        }

        /// <summary>
        /// 下载未完成和错误的Project模板
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadProjects(bool? result)
        {
            List<Project> list;
            if (result.HasValue)
            {
                list = db.Projects.Where(e => e.Result != result.Value).ToList();
            }
            else
            {
                list = db.Projects.Where(e => e.Result != null).ToList();
            }

            var workbook = XslHelper.GetWorkbook("templates/modelSelf.xlsx");

            var sheet = workbook.GetSheetAt(0);

            int y = 1;
            foreach (var item in list)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    row = sheet.CreateRow(y);
                ICell cell = row.GetCell(0);
                if (cell == null)
                    cell = row.CreateCell(0);
                cell.SetCellValue(y);

                ICell cell2 = row.GetCell(1);
                if (cell2 == null)
                    cell2 = row.CreateCell(1);
                cell2.SetCellValue(City.浙江省.ToString() + "," + item.City.ToString() + "," + item.County);

                ICell cell3 = row.GetCell(2);
                if (cell3 == null)
                    cell3 = row.CreateCell(2);
                cell3.SetCellValue(item.ID);
                ICell cell4 = row.GetCell(3);
                if (cell4 == null)
                    cell4 = row.CreateCell(3);
                cell4.SetCellValue(item.Name);
                y++;
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "自检表.xlsx");
        }

        /// <summary>
        /// 上传一部分项目，验证并更新到Project
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects(FormCollection form)
        {
            var file = UploadHelper.GetPostedFile(HttpContext);

            var filePath = UploadHelper.UploadExcel(file);

            var uploadFile = new UploadFile
            {
                City = CurrentUser.City,
                CreateTime = DateTime.Now,
                FileName = file.FileName,
                SavePath = filePath
            };

            db.Files.Add(uploadFile);

            db.SaveChanges();
            //上传成功后跳转到check页面进行检查，参数是File的ID
            return RedirectToAction("Check", new { id = uploadFile.ID });
        }

        public ActionResult Check(int id)
        {
            var file = db.Files.FirstOrDefault(e => e.ID == id);
            if (file == null)
            {
                throw new ArgumentException("参数错误");
            }

            var filePath = UploadHelper.GetAbsolutePath(file.SavePath);
            //读取文件进行检查
            Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();
            Dictionary<string, int> ship = new Dictionary<string, int>();
            DetectEngine Engine = new DetectEngine(filePath);
            string fault = "";
            if (!Engine.CheckExcel(filePath, ref fault, ref Error, ref ship))
            {
                throw new ArgumentException("检索失败");
            }
            //检查完毕，更新Projects
            var projects = db.Projects.Where(x => x.City == CurrentUser.City);
            foreach (var item in projects)
            {
                if (ship.ContainsKey(item.ID))
                {
                    if (Error.ContainsKey(item.ID))
                    {
                        item.Note = "";
                        item.Result = false;
                        foreach (var Message in Error[item.ID])
                        {
                            item.Note += Message + "；";
                        }
                    }
                    else
                    {
                        item.Result = true;
                        item.Note = "";
                    }
                    db.SaveChanges();
                }
            }

            return RedirectToAction("Index", new { result = false });
        }

        /// <summary>
        /// 下载表2
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex2()
        {
            var workbook = XslHelper.GetWorkbook("templates/表2.xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认无问题项目清单");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int rowIndex = 6;
            int rowNumber = 1;

            var textCellStyle = workbook.GetCellStyle(XslHeaderStyle.文本);

            foreach (var item in list)
            {
                var row = CreateRow(sheet, rowIndex, 0, 7, textCellStyle);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);
                row.Cells[3].SetCellValue(item.ID);
                row.Cells[4].SetCellValue(item.Name);

                rowIndex++;
                rowNumber++;
            }

            return GetFileResult(workbook, "附表2.xls");
        }

        /// <summary>
        /// 下载表3
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex3()
        {
            var workbook = XslHelper.GetWorkbook("templates/表3.xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点项目复核确认总表");
            var textCellStyle = workbook.GetCellStyle(XslHeaderStyle.文本);


            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int rowIndex = 6;
            int rowNumber = 1;
            foreach (var item in list)
            {
                var row = CreateRow(sheet, rowIndex, 0, 9, textCellStyle);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);
                row.Cells[3].SetCellValue(item.ID);
                row.Cells[4].SetCellValue(item.Name);

                rowIndex++;
                rowNumber++;
            }

            return GetFileResult(workbook, "附表3.xls");
        }

        /// <summary>
        /// 下载表4
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex4()
        {
            var workbook = XslHelper.GetWorkbook("templates/表4.xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目申请删除项目清单");

            var textCellStyle = workbook.GetCellStyle(XslHeaderStyle.文本);


            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            var rowIndex = 6;
            var rowNumber = 1;
            foreach (var item in list)
            {
                var row = CreateRow(sheet, rowIndex, 0, 6, textCellStyle);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);
                row.Cells[3].SetCellValue(item.ID);

                rowIndex++;
                rowNumber++;
            }

            return GetFileResult(workbook, "附表4.xls");
        }

        /// <summary>
        /// 下载表5
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex5()
        {
            var workbook = XslHelper.GetWorkbook("templates/表5.xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目备案信息错误项目清单");

            var textCellStyle = workbook.GetCellStyle(XslHeaderStyle.文本);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            var rowIndex = 5;
            var rowNumber = 1;
            foreach (var item in list)
            {
                var row = CreateRow(sheet, rowIndex, 0, 8, textCellStyle);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);
                row.Cells[3].SetCellValue(item.ID);
                row.Cells[4].SetCellValue(item.Name);
                row.Cells[8].SetCellValue("是");
                rowIndex++;
                rowNumber++;
            }

            return GetFileResult(workbook, "附表5.xls");
        }

        /// <summary>
        /// 下载表6
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex6()
        {
            var workbook = XslHelper.GetWorkbook("templates/表6.xlsx");
            var sheet = workbook.GetSheetAt(0);

            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目设计二调新增耕地项目清单");

            return GetFileResult(workbook, "附表6.xls");
        }

        /// <summary>
        /// 下载表7
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex7()
        {
            var workbook = XslHelper.GetWorkbook("templates/表7.xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目耕地质量等别修改项目清单");

            var textCellStyle = workbook.GetCellStyle(XslHeaderStyle.文本);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            var rowIndex = 5;
            var rowNumber = 1;
            foreach (var item in list)
            {
                var row = CreateRow(sheet, rowIndex, 0, 8, textCellStyle);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);
                row.Cells[3].SetCellValue(item.ID);
                row.Cells[4].SetCellValue(item.Name);
                row.Cells[8].SetCellValue("是");
                rowIndex++;
                rowNumber++;
            }

            return GetFileResult(workbook, "附表7.xls");
        }

        /// <summary>
        /// 下载表8
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex8()
        {
            var workbook = XslHelper.GetWorkbook("templates/表8.xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目占补平衡指标核减项目清单");

            var textCellStyle = workbook.GetCellStyle(XslHeaderStyle.文本);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            var rowIndex = 5;
            var rowNumber = 1;
            foreach (var item in list)
            {
                var row = CreateRow(sheet, rowIndex, 0, 8, textCellStyle);
                row.Cells[0].SetCellValue(rowNumber.ToString());
                row.Cells[1].SetCellValue(item.City.ToString());
                row.Cells[2].SetCellValue(item.County);
                row.Cells[3].SetCellValue(item.ID);
                row.Cells[4].SetCellValue(item.Name);

                rowIndex++;
                rowNumber++;
            }

            return GetFileResult(workbook, "附表8.xls");
        }

        /// <summary>
        /// 下载表9
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex9()
        {
            var workbook = XslHelper.GetWorkbook("templates/表9.xlsx");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目新增耕地二级地类与耕地质量等别确认表");

            var textCellStyle = workbook.GetCellStyle(XslHeaderStyle.文本);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            var rowIndex = 6;
            var rowNumber = 1;
            var dkNames = new[] { "水田", "水浇地", "旱地" };
            foreach (var item in list)
            {
                
                foreach (var name in dkNames)
                {
                    var row = CreateRow(sheet, rowIndex, 0, 22, textCellStyle);
                    row.Cells[6].SetCellValue(dkNames[rowIndex - 6]);
                    for (int i = 7; i < 22; i++)
                    {
                        row.Cells[i].SetCellValue("（亩）");
                    }
                    row.Cells[22].SetCellValue("是");
                    rowIndex++;
                }

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
                
                rowNumber++;
            }

            return GetFileResult(workbook, "附表9.xls");
        }

        private IRow CreateRow(ISheet sheet, int rowIndex, int startColumnNumber, int lastColumnNumber, ICellStyle style)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            for (var i = startColumnNumber; i <= lastColumnNumber; i++)
            {
                var cell = row.GetCell(i);
                if (cell == null)
                {
                    cell = row.CreateCell(i);
                    cell.CellStyle = style;
                }
            }
            return row;
        }

        private ActionResult GetFileResult(IWorkbook workbook, string fileName)
        {
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                return File(ms.ToArray(), "application/ms-excel", fileName);
            }
        }
    }
}
