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

            string checkSelfExcel = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "ModelSelf.xlsx");
            IWorkbook workbook;
            try
            {
                FileStream fs = new FileStream(checkSelfExcel, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
                fs.Close();
            }
            catch
            {
                throw new ArgumentException("打开自查表模板失败");
            }

            ISheet sheet = workbook.GetSheetAt(0);
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
        /// 下载表格时候，使用到表格的格式
        /// </summary>
        public enum stylexls
        {
            大头,
            小头,
            小小头,
            文本,
            默认
        }

        public static ICellStyle GetCellStyle(IWorkbook workbook, stylexls str)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            IFont fontBigheader = workbook.CreateFont();
            fontBigheader.FontHeightInPoints = 22;
            fontBigheader.FontName = "微软雅黑";
            fontBigheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            IFont fontSmallheader = workbook.CreateFont();
            fontSmallheader.FontHeightInPoints = 14;
            fontSmallheader.FontName = "黑体";

            IFont fontText = workbook.CreateFont();
            fontText.FontHeightInPoints = 12;
            fontText.FontName = "宋体";

            IFont fontthinheader = workbook.CreateFont();
            fontthinheader.FontName = "宋体";
            fontthinheader.FontHeightInPoints = 11;
            fontthinheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;


            IFont font1 = workbook.CreateFont();
            font1.FontName = "宋体";
            font1.FontHeightInPoints = 9;


            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            //边框颜色
            cellStyle.BottomBorderColor = HSSFColor.OliveGreen.Black.Index;
            cellStyle.TopBorderColor = HSSFColor.OliveGreen.Black.Index;

            //背景图形
            cellStyle.FillForegroundColor = HSSFColor.White.Index;
            cellStyle.FillBackgroundColor = HSSFColor.Black.Index;

            //文本对齐  左对齐  居中  右对齐  现在是居中
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            //垂直对齐
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            //自动换行
            cellStyle.WrapText = true;

            //缩进
            cellStyle.Indention = 0;

            switch (str)
            {
                case stylexls.大头:
                    cellStyle.SetFont(fontBigheader);
                    break;
                case stylexls.小头:
                    cellStyle.SetFont(fontSmallheader);
                    break;
                case stylexls.默认:
                    cellStyle.SetFont(fontText);
                    break;
                case stylexls.小小头:
                    cellStyle.SetFont(fontthinheader);
                    break;

                case stylexls.文本:
                    cellStyle.SetFont(font1);
                    break;
            }


            return cellStyle;
        }


        /// <summary>
        /// 下载表2
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex2()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表2.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow HeadRow = sheet.GetRow(1);
            ICell HeadCell = HeadRow.GetCell(0);
            string Header = CurrentUser.City.ToString() + "重点复核确认无问题项目清单";
            HeadCell.SetCellValue(Header);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int y = 6;
            int s = 1;
            foreach (var item in list)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    row = sheet.CreateRow(y);
                y++;


                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(s.ToString());
                s++;

                cell = row.GetCell(1);
                if (cell == null)
                {
                    cell = row.CreateCell(1);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.City.ToString());

                cell = row.GetCell(2);
                if (cell == null)
                {
                    cell = row.CreateCell(2);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.County);

                cell = row.GetCell(3);
                if (cell == null)
                {
                    cell = row.CreateCell(3);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.ID.ToString());

                cell = row.GetCell(4);
                if (cell == null)
                {
                    cell = row.CreateCell(4);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.Name);

                cell = row.GetCell(5);
                if (cell == null)
                {
                    cell = row.CreateCell(5);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(6);
                if (cell == null)
                {
                    cell = row.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(7);
                if (cell == null)
                {
                    cell = row.CreateCell(7);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表2.xls");
        }

        /// <summary>
        /// 下载表3
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex3()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表3.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow HeadRow = sheet.GetRow(1);
            ICell HeadCell = HeadRow.GetCell(0);
            string Header = CurrentUser.City.ToString() + "重点项目复核确认总表";
            HeadCell.SetCellValue(Header);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int y = 6;
            int s = 1;
            foreach (var item in list)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    row = sheet.CreateRow(y);
                y++;


                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(s.ToString());
                s++;

                cell = row.GetCell(1);
                if (cell == null)
                {
                    cell = row.CreateCell(1);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.City.ToString());

                cell = row.GetCell(2);
                if (cell == null)
                {
                    cell = row.CreateCell(2);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.County);

                cell = row.GetCell(3);
                if (cell == null)
                {
                    cell = row.CreateCell(3);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.ID.ToString());

                cell = row.GetCell(4);
                if (cell == null)
                {
                    cell = row.CreateCell(4);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.Name);

                cell = row.GetCell(5);
                if (cell == null)
                {
                    cell = row.CreateCell(5);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(6);
                if (cell == null)
                {
                    cell = row.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(7);
                if (cell == null)
                {
                    cell = row.CreateCell(7);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row.GetCell(8);
                if (cell == null)
                {
                    cell = row.CreateCell(8);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row.GetCell(9);
                if (cell == null)
                {
                    cell = row.CreateCell(9);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表2.xls");
        }

        /// <summary>
        /// 下载表4
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex4()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表4.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow HeadRow = sheet.GetRow(1);
            ICell HeadCell = HeadRow.GetCell(0);
            string Header = CurrentUser.City.ToString() + "重点复核确认项目申请删除项目清单";
            HeadCell.SetCellValue(Header);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int y = 6;
            int s = 1;
            foreach (var item in list)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    row = sheet.CreateRow(y);
                y++;


                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(s.ToString());
                s++;

                cell = row.GetCell(1);
                if (cell == null)
                {
                    cell = row.CreateCell(1);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.City.ToString());

                cell = row.GetCell(2);
                if (cell == null)
                {
                    cell = row.CreateCell(2);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.County);

                cell = row.GetCell(3);
                if (cell == null)
                {
                    cell = row.CreateCell(3);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.ID.ToString());

                cell = row.GetCell(4);
                if (cell == null)
                {
                    cell = row.CreateCell(4);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(5);
                if (cell == null)
                {
                    cell = row.CreateCell(5);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(6);
                if (cell == null)
                {
                    cell = row.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表4.xls");
        }

        /// <summary>
        /// 下载表5
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex5()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表5.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow HeadRow = sheet.GetRow(1);
            ICell HeadCell = HeadRow.GetCell(0);
            string Header = CurrentUser.City.ToString() + "重点复核确认项目备案信息错误项目清单";
            HeadCell.SetCellValue(Header);

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int y = 5;
            int s = 1;
            foreach (var item in list)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    row = sheet.CreateRow(y);
                y++;


                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(s.ToString());
                s++;

                cell = row.GetCell(1);
                if (cell == null)
                {
                    cell = row.CreateCell(1);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.City.ToString());

                cell = row.GetCell(2);
                if (cell == null)
                {
                    cell = row.CreateCell(2);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.County);

                cell = row.GetCell(3);
                if (cell == null)
                {
                    cell = row.CreateCell(3);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.ID.ToString());

                cell = row.GetCell(4);
                if (cell == null)
                {
                    cell = row.CreateCell(4);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.Name);


                cell = row.GetCell(5);
                if (cell == null)
                {
                    cell = row.CreateCell(5);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(6);
                if (cell == null)
                {
                    cell = row.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(7);
                if (cell == null)
                {
                    cell = row.CreateCell(7);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row.GetCell(8);
                if (cell == null)
                {
                    cell = row.CreateCell(8);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("是");

            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表5.xls");
        }

        /// <summary>
        /// 下载表6
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex6()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表6.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow HeadRow = sheet.GetRow(1);
            ICell HeadCell = HeadRow.GetCell(0);
            string Header = CurrentUser.City.ToString() + "重点复核确认项目设计二调新增耕地项目清单";
            HeadCell.SetCellValue(Header);
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表6.xls");
        }

        /// <summary>
        /// 下载表7
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex7()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表7.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow headRow = sheet.GetRow(1);
            ICell headCell = headRow.GetCell(0);
            string Header = CurrentUser.City.ToString() + "重点复核确认项目耕地质量等别修改项目清单";
            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int y = 5;
            int s = 1;
            foreach (var item in list)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    row = sheet.CreateRow(y);
                y++;


                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(s.ToString());
                s++;

                cell = row.GetCell(1);
                if (cell == null)
                {
                    cell = row.CreateCell(1);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.City.ToString());

                cell = row.GetCell(2);
                if (cell == null)
                {
                    cell = row.CreateCell(2);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.County);

                cell = row.GetCell(3);
                if (cell == null)
                {
                    cell = row.CreateCell(3);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(4);
                if (cell == null)
                {
                    cell = row.CreateCell(4);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }



                cell = row.GetCell(5);
                if (cell == null)
                {
                    cell = row.CreateCell(5);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(6);
                if (cell == null)
                {
                    cell = row.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(7);
                if (cell == null)
                {
                    cell = row.CreateCell(7);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row.GetCell(8);
                if (cell == null)
                {
                    cell = row.CreateCell(8);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("是");

            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表7.xls");
        }

        /// <summary>
        /// 下载表8
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex8()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表8.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow headRow = sheet.GetRow(1);
            ICell headCell = headRow.GetCell(0);
            string Header = CurrentUser.City.ToString() + "重点复核确认项目占补平衡指标核减项目清单";
            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int y = 5;
            int s = 1;
            foreach (var item in list)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    row = sheet.CreateRow(y);
                y++;


                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(s.ToString());
                s++;

                cell = row.GetCell(1);
                if (cell == null)
                {
                    cell = row.CreateCell(1);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.City.ToString());

                cell = row.GetCell(2);
                if (cell == null)
                {
                    cell = row.CreateCell(2);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.County);

                cell = row.GetCell(3);
                if (cell == null)
                {
                    cell = row.CreateCell(3);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(4);
                if (cell == null)
                {
                    cell = row.CreateCell(4);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }



                cell = row.GetCell(5);
                if (cell == null)
                {
                    cell = row.CreateCell(5);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(6);
                if (cell == null)
                {
                    cell = row.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }


                cell = row.GetCell(7);
                if (cell == null)
                {
                    cell = row.CreateCell(7);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row.GetCell(8);
                if (cell == null)
                {
                    cell = row.CreateCell(8);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表8.xls");
        }

        /// <summary>
        /// 下载表9
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex9()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "表9.xls");
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow headRow = sheet.GetRow(1);
            ICell headCell = headRow.GetCell(0);


            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int y = 6;
            int s = 1;
            foreach (var item in list)
            {
                IRow row1 = sheet.GetRow(y);
                if (row1 == null)
                    row1 = sheet.CreateRow(y);
                y++;
                ICell cell = row1.GetCell(0);
                if (cell == null)
                {
                    cell = row1.CreateCell(0);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(s.ToString());
                s++;

                cell = row1.GetCell(1);
                if (cell == null)
                {
                    cell = row1.CreateCell(1);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.City.ToString());

                cell = row1.GetCell(2);
                if (cell == null)
                {
                    cell = row1.CreateCell(2);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue(item.County);

                cell = row1.GetCell(3);
                if (cell == null)
                {
                    cell = row1.CreateCell(3);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row1.GetCell(4);
                if (cell == null)
                {
                    cell = row1.CreateCell(4);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row1.GetCell(5);
                if (cell == null)
                {
                    cell = row1.CreateCell(5);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }

                cell = row1.GetCell(6);
                if (cell == null)
                {
                    cell = row1.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("水田");

                for (int i = 7; i < 22; i++)
                {
                    cell = row1.GetCell(i);
                    if (cell == null)
                    {
                        cell = row1.CreateCell(i);
                        cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                    }
                    cell.SetCellValue("    (亩)");
                }

                cell = row1.GetCell(22);
                if (cell == null)
                {
                    cell = row1.CreateCell(22);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("是");





                IRow row2 = sheet.GetRow(y);
                if (row2 == null)
                    row2 = sheet.CreateRow(y);
                y++;

                cell = row2.GetCell(6);
                if (cell == null)
                {
                    cell = row2.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("水浇地");

                for (int i = 7; i < 22; i++)
                {
                    cell = row2.GetCell(i);
                    if (cell == null)
                    {
                        cell = row2.CreateCell(i);
                        cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                    }
                    cell.SetCellValue("    (亩)");
                }

                cell = row2.GetCell(22);
                if (cell == null)
                {
                    cell = row2.CreateCell(22);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("是");


                IRow row3 = sheet.GetRow(y);
                if (row3 == null)
                    row3 = sheet.CreateRow(y);
                y++;

                cell = row3.GetCell(6);
                if (cell == null)
                {
                    cell = row3.CreateCell(6);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("旱地");

                for (int i = 7; i < 22; i++)
                {
                    cell = row3.GetCell(i);
                    if (cell == null)
                    {
                        cell = row3.CreateCell(i);
                        cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                    }
                    cell.SetCellValue("    (亩)");
                }

                cell = row3.GetCell(22);
                if (cell == null)
                {
                    cell = row3.CreateCell(22);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.文本);
                }
                cell.SetCellValue("是");

                /*
                 * 合并前6列的单元格
                 */
                for (int j = 0; j < 6; j++)
                {
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress((y - 3), (y - 1), j, j));
                }

            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "附表9.xls");
        }
    }
}
