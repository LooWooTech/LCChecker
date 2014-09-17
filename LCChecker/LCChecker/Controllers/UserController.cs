using LCChecker.Models;
using NPOI.HSSF.UserModel;
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
        public ActionResult Index()
        {

            ViewBag.ProjectSummary = new Summary
            {
                City = CurrentUser.City,
                TotalCount = db.Projects.Count(e => e.City == CurrentUser.City),
                SuccessCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == true),
                ErrorCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == false),

            };
            ViewBag.ReportSummary = new Summary
            {
                City = CurrentUser.City,
                TotalCount = db.Reports.Count(e => e.City == CurrentUser.City),
                SuccessCount = db.Reports.Count(e => e.City == CurrentUser.City && e.Result == true),
                ErrorCount = db.Reports.Count(e => e.City == CurrentUser.City && e.Result == false),
            };
            return View();
        }

        public ActionResult Projects(ResultFilter result = ResultFilter.All, int page = 1)
        {
            var paging = new Page(page);
            ViewBag.List = ProjectHelper.GetProjects(CurrentUser.City, result, paging);
            ViewBag.Page = paging;
            return View();
        }

        public ActionResult Coordinates()
        {
            return View();
        }

        public ActionResult UploadCoordinates()
        {
            throw new NotImplementedException();
        }

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

        /// <summary>
        /// 上传一部分项目，验证并更新到Project
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects(int id = 0)
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

            var summary = new Summary
            {
                City = CurrentUser.City,
                TotalCount = db.Projects.Count(e => e.City == CurrentUser.City),
                SuccessCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == true),
                ErrorCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == false),

            };
            if (id != 0)
            {
                string MastPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/", CurrentUser.City.ToString() + ".xls");
                filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePath);
                //CheckReport4 engine = new CheckReport4(MastPath);
                //string fault = "";
                //if (!engine.Check(filePath, ref fault))
                //{
                //    //验证失败;
                //}

                CheckReport5 engine = new CheckReport5(MastPath);
                string fault = "";
                if (!engine.Check(filePath, ref fault))
                { 
                    
                }

                ReportType reportType=(ReportType)Enum.Parse(typeof(ReportType),id.ToString());
                
                var record = db.Reports.Where(e => e.City == CurrentUser.City && e.Type ==reportType).FirstOrDefault();
                if (record != null)
                {
                    record.Result = false;
                    db.Entry(record).State = EntityState.Modified;
                    db.SaveChanges();
                }
                return RedirectToAction("Reports");
            }
            //全部结束，进入第二阶段
            /*if (summary.TotalCount > 0 && summary.TotalCount == summary.SuccessCount)
            {
                return RedirectToAction("CheckIndex", new { id = uploadFile.ID });
            }*/
            //上传成功后跳转到check页面进行检查，参数是File的ID
            return RedirectToAction("Check", new { id = uploadFile.ID, TypeId = id });
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
            var errors = new Dictionary<string, List<string>>();
            var ships = new Dictionary<string, int>();
            var detectEngine = new DetectEngine(filePath);
            var fault = "";
            if (!detectEngine.CheckExcel(filePath, ref fault, ref errors, ref ships))
            {
                throw new ArgumentException("检索失败：" + fault);
            }

            var masterfile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/", CurrentUser.City.ToString() + ".xls");
            //string Masterfile = @"D:\Work\浙江省土地整治项目核查平台\0914\trunk\LCChecker\LCChecker\App_Data\湖州市.xls";
            var list = db.Projects.Where(x => x.City == CurrentUser.City).ToList();
            
            if (!detectEngine.SaveCurrent(filePath, masterfile, ref fault, errors, ships,list))
            {
                throw new ArgumentException("保存正确项目失败");
            }

            //检查完毕，更新Projects
            var projects = db.Projects.Where(e => e.City == CurrentUser.City);
            foreach (var item in projects)
            {
                if (ships.ContainsKey(item.ID))
                {
                    if (errors.ContainsKey(item.ID))
                    {
                        item.Note = "";
                        item.Result = false;
                        var errs = errors[item.ID];
                        var i = 1;
                        foreach (var msg in errs)
                        {
                            item.Note += string.Format("({0}){1}；", i, msg); ;
                            i++;
                        }
                    }
                    else
                    {
                        item.Result = true;
                        item.Note = "";
                    }
                }
            }

            db.SaveChanges();

            return RedirectToAction("projects", new { result = (int)ResultFilter.Error });
        }

        public ActionResult CheckIndex(int id, ReportType typeId)
        {
            var file = db.Files.FirstOrDefault(e => e.ID == id);
            if (file == null)
            {
                throw new ArgumentException("参数错误");
            }

            var filePath = UploadHelper.GetAbsolutePath(file.SavePath);

            string masterPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", CurrentUser.City.ToString() + ".xls");
            //string fault = "";
//<<<<<<< .mine
//            CheckReport2 Engine = new CheckReport2(masterPath);
//            if (!Engine.Check(filePath, ref fault))
//            {
//                throw new ArgumentException("检索附表失败"+fault);
//            }
//=======
            //CheckReport2 Engine = new CheckReport2(masterPath);
            //if (!Engine.Check(filePath, ref fault))
            //{
            //    throw new ArgumentException("检索附表失败");
            //}
//>>>>>>> .r93

            //if (Engine.Error.Count() != 0)
            //{

            //}

            //var message = db.Reports.Where(x => x.City == CurrentUser.City && x.Type == typeId).FirstOrDefault();
            //if (message == null)
            //{
            //    throw new ArgumentException("未找到上传信息");
            //}
            //message.Result = true;
            //db.Entry(message).State = EntityState.Modified;
            //db.SaveChanges();



            Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();
            ViewBag.Error = Error;
            return View(Error);
        }

        /// <summary>
        /// 下载未完成和错误的Project模板
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadProjects(bool? result)
        {
            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            var workbook = XslHelper.GetWorkbook("templates/modelSelf.xlsx");

            var sheet = workbook.GetSheetAt(0);
            var rowIndex = 1;

            sheet.InsertRow(rowIndex, list.Count - 1);

            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
                row.Cells[0].SetCellValue(rowIndex);
                row.Cells[1].SetCellValue(City.浙江省.ToString() + "," + item.City.ToString() + "," + item.County);
                row.Cells[2].SetCellValue(item.ID);
                row.Cells[3].SetCellValue(item.Name);

                rowIndex++;
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "自检表.xlsx");
        }

        /// <summary>
        /// 下载表2
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex2()
        {
            var workbook = XslHelper.GetWorkbook("Templates/2.xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认无问题项目清单");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int count = list.Count();
            int rowIndex = 6;
            int rowNumber = 1;

            sheet.InsertRow(rowIndex, count - 1);

            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
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
            var workbook = XslHelper.GetWorkbook("Templates/3.xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点项目复核确认总表");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int count = list.Count();
            int rowIndex = 6;
            int rowNumber = 1;
            sheet.InsertRow(rowIndex, count - 1);

            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
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
            var workbook = XslHelper.GetWorkbook("Templates/4.xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目申请删除项目清单");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            int count = list.Count();
            var rowIndex = 6;
            var rowNumber = 1;
            sheet.InsertRow(rowIndex, count - 1);

            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
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
            var workbook = XslHelper.GetWorkbook("Templates/5.xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目备案信息错误项目清单");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            var count = list.Count();
            var rowIndex = 5;
            var rowNumber = 1;
            sheet.InsertRow(rowIndex, count - 1);

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

            return GetFileResult(workbook, "附表5.xls");
        }

        /// <summary>
        /// 下载表6
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex6()
        {
            var workbook = XslHelper.GetWorkbook("Templates/6.xls");

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
            var workbook = XslHelper.GetWorkbook("Templates/7.xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目耕地质量等别修改项目清单");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            var count = list.Count();
            var rowIndex = 5;
            var rowNumber = 1;
            sheet.InsertRow(rowIndex, count - 1);

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

            return GetFileResult(workbook, "附表7.xls");
        }

        /// <summary>
        /// 下载表8
        /// </summary>
        /// <returns></returns>
        public ActionResult DownReportExcelIndex8()
        {
            var workbook = XslHelper.GetWorkbook("Templates/8.xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目占补平衡指标核减项目清单");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            var count = list.Count();
            var rowIndex = 5;
            var rowNumber = 1;
            sheet.InsertRow(rowIndex, count - 1);

            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
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
            var workbook = XslHelper.GetWorkbook("Templates/9.xls");
            var sheet = workbook.GetSheetAt(0);
            sheet.GetRow(1).Cells[0].SetCellValue(CurrentUser.City.ToString() + "重点复核确认项目新增耕地二级地类与耕地质量等别确认表");

            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
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

            return GetFileResult(workbook, "附表9.xls");
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
