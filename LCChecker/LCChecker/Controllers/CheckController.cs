using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    public class CheckController : BaseController
    {
        private LCDbContext db = new LCDbContext();
        //
        // GET: /Check/

        public ActionResult Index()
        {
            //ceshi();
            return View();
        }

        [HttpPost]
        public ActionResult Index(User user)
        {
            User logOne = db.USER.Where(x => x.logName == user.logName && x.password == user.password).FirstOrDefault();
            if (logOne == null)
            {
                return View();
            }
            Session["id"] = logOne.id;
            Session["name"] = logOne.name;
            if (logOne.flag)
            {
                return RedirectToAction("Admin");
            }
            return RedirectToAction("Region", new { regionName=logOne.name});
        }


        /*管理员*/
        public ActionResult Admin()
        {
            return View(db.DETECT);
        }

        /*区域用户登录*/
        public ActionResult Region(string regionName)
        {
            Detect record = db.DETECT.Where(x => x.region == regionName).FirstOrDefault();
            if (record == null)
            {
                return RedirectToAction("Index");
            }
            ViewBag.submits = record.submit;
            ViewBag.name = regionName;
            if (record.flag)
            {
                return View("success");
            }
            else {
                ViewBag.records = db.SUBRECORD.Where(x => x.regionId==record.Id).ToList();
                return View();
            }
        }

        /*用户上传文件*/
        [HttpPost]
        public ActionResult FileUpload(FormCollection form)
        {
            if (Session["id"] == null||Session["name"]==null)
            {
                return RedirectToAction("Index");
            }
            string name=Session["name"].ToString();
            if (Request.Files.Count == 0)
            {
                ViewBag.ErrorMessage = "请选择文件上传";
                return View();
            }

            HttpPostedFileBase file = Request.Files[0];
            string ext=Path.GetExtension(file.FileName);
            if (ext != ".xls" && ext != ".xlsx")
            {
                string error= "你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格";
                Session["Error"] = error;
                return RedirectToAction("Region", new { regionName = name });
            }

            if (file.ContentLength == 0||file.ContentLength>20971520)
            {
                Session["Error"] = "你上传的文件0字节或者大于20M 无法读取表格";
                return RedirectToAction("Region", new { regionName = name });
            }
            else {
                Detect record = db.DETECT.Where(x => x.region == name).FirstOrDefault();
                string filePath = null;
                if (record == null)
                {
                    record = new Detect() {  region = name,submit = 1};
                    if (ModelState.IsValid)
                    {
                        db.DETECT.Add(record);
                        db.SaveChanges();
                    }
                }
                else {
                    record.submit++;
                    db.Entry(record).State = EntityState.Modified;
                    db.SaveChanges();
                    
                }
                if (ext == ".xls")
                {
                    filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + name + "/" + record.submit), "NO" + record.submit + ".xls");
                }
                else {
                    filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + name + "/" + record.submit), "NO" + record.submit + ".xlsx");
                }
                
                string Catalogue = HttpContext.Server.MapPath("../Uploads/" + name+"/"+record.submit);
                if (!Directory.Exists(Catalogue))
                {
                    try
                    {
                        Directory.CreateDirectory(Catalogue);
                    }
                    catch(Exception er)
                    {
                        Session["Error"] = "错误：服务器创建目录失败,错误信息："+er.Message;
                        return RedirectToAction("Region", new { regionName = name });
                    }  
                }
                
                try
                {
                    FileStream fs = new FileStream(filePath, FileMode.Create);
                    fs.Close();
                }
                catch {
                    Session["Error"] ="错误：服务器保存上传文件失败";
                    return RedirectToAction("Region", new { regionName = name });
                } 
                file.SaveAs(filePath);
                SubRecord submit = new SubRecord();
                submit.Format = ext;
                submit.regionId = record.Id;
                submit.name = file.FileName;
                submit.submits = record.submit;
                if (ModelState.IsValid)
                {
                    db.SUBRECORD.Add(submit);
                    db.SaveChanges();
                }
                return RedirectToAction("Check", "Check", new { region = name, SubmitFile=filePath});
            }
        }


        /*
         * 下载表格
         * 根据fileType来判断下载那种表格 fileType为sub 提交错误表  fileType 为sum 下载总表错误 
         * region 指定那个区域 
         */
        public ActionResult DownExcel(string fileType,string region)
        {
            Detect Area = db.DETECT.Where(x => x.region == region).FirstOrDefault();
            if (Area == null)
            {
                return RedirectToAction("Index");
            }
            string DownFilePath ;
            string fileName;
            if (fileType == "sub")
            {
                fileName = "提交错误表.xlsx";
                DownFilePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + Area.submit), "subErrorExcel.xlsx");
            }
            else if (fileType == "sum")
            {
                fileName = "总错误表.xlsx";
                DownFilePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + Area.submit), "summaryErrorExcel.xlsx");
            }
            else {
                fileName = region + ".xlsx";
                DownFilePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region), "summary.xlsx");
            }
            FileStream fs;
            try { 
                fs = new FileStream(DownFilePath, FileMode.Open, FileAccess.Read);
            }
            catch(Exception er)
            {
                Session["Error"] = "该用户可能还没有提交表格，错误原因：" + er.Message;
                return RedirectToAction("admin");
            }
            byte[] fileContents = new byte[(int)fs.Length];
            fs.Read(fileContents, 0, fileContents.Length);
            fs.Close();
            return File(fileContents, "application/ms-excel", fileName);
        }

        /*下载提交表格
         */
        public ActionResult DownSubExcel(string region, string times,string Type,string fileName)
        {
            string SourcePath = "NO" + times+Type;
            string FilePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + times),SourcePath);
            FileStream fs ;
            try {
                fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
            }
            catch {
                return View();
            }
            byte[] fileContents = new byte[(int)fs.Length];
            fs.Read(fileContents, 0, fileContents.Length);
            fs.Close();
            return File(fileContents, "application/ms-excel", fileName);
            
        }
        

        /*
         * 管理员下载区域检查情况
         */
        public FileResult DownLoad()
        {
            MemoryStream ms = new MemoryStream();
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row = sheet.CreateRow(0);
            row.Height = 20 * 20;
            for (int x= 0; x < 8; x++)
            {
                sheet.SetColumnWidth(x, 20 * 256);
            }
                
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 7));
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 2, 2, 5));
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 3, 0, 0));
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 3, 1, 1));
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 3, 6, 6));
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 3, 7, 7));
            
            ICell Cell = row.CreateCell(0);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.大头);
            Cell.SetCellValue("2014年农村土地整治检测监管系统清查检查表");
            for (int x = 1; x < 8; x++)
            {
                ICell xCell = row.CreateCell(x);
                Cell.CellStyle = GetCellStyle(workbook, stylexls.大头);
            }
            IRow row2 = sheet.CreateRow(1);
            for (int x = 0; x < 8; x++)
            {
                ICell xCell = row2.CreateCell(x);
                Cell.CellStyle = GetCellStyle(workbook, stylexls.大头);
            }


                
            row = sheet.CreateRow(2);

            Cell = row.CreateCell(0);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("序号");

            Cell = row.CreateCell(1);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("行政单位");

            Cell = row.CreateCell(2);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("下发清单情况");

            for (int x = 3; x < 6; x++)
            {
                Cell = row.CreateCell(x);
                Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            }

            Cell = row.CreateCell(6);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("提交次数");

            Cell = row.CreateCell(7);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("目前通过个数");


            row = sheet.CreateRow(3);

            for (int x = 0; x < 2; x++)
            {
                Cell = row.CreateCell(x);
                Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            }
             
            Cell = row.CreateCell(2);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("项目个数");

            Cell = row.CreateCell(3);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("总规模");

            Cell = row.CreateCell(4);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("新增耕地面积");

            Cell = row.CreateCell(5);
            Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            Cell.SetCellValue("可用于占补平衡面积");

            for (int x = 6; x < 8; x++)
            {
                Cell = row.CreateCell(x);
                Cell.CellStyle = GetCellStyle(workbook, stylexls.小头);
            }

            List<Detect> Detects = db.DETECT.ToList();
            int i=4;
            foreach (var item in Detects)
            {
                IRow newRow = sheet.CreateRow(i++);
                ICell newCell = newRow.CreateCell(0);
                newCell.CellStyle = GetCellStyle(workbook,stylexls.默认);
                newCell.SetCellValue(item.Id);
                newCell = newRow.CreateCell(1);
                newCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                newCell.SetCellValue(item.region);
                newCell = newRow.CreateCell(2);
                newCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                newCell.SetCellValue(item.sum);
                newCell = newRow.CreateCell(3);
                newCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                newCell.SetCellValue(item.totalScale);
                newCell = newRow.CreateCell(4);
                newCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                newCell.SetCellValue(item.AddArea);
                newCell = newRow.CreateCell(5);
                newCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                newCell.SetCellValue(item.available);
                newCell = newRow.CreateCell(6);
                newCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                newCell.SetCellValue(item.submit);
                newCell = newRow.CreateCell(7);
                newCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                newCell.SetCellValue(item.Correct);
            }
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i, 0, 1));
            IRow hRow = sheet.CreateRow(i);
            ICell hCell = hRow.CreateCell(0);
            hCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
            hCell.SetCellValue("合计");
            hCell=hRow.CreateCell(1);
            hCell.CellStyle=GetCellStyle(workbook,stylexls.默认);

            int sum=db.DETECT.Sum(x=>x.sum);
            hCell = hRow.CreateCell(2);
            hCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
            hCell.SetCellValue(sum);

            sum = db.DETECT.Sum(x => x.totalScale);
            hCell = hRow.CreateCell(3);
            hCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
            hCell.SetCellValue(sum);
            
            sum = db.DETECT.Sum(x => x.AddArea);
            hCell = hRow.CreateCell(4);
            hCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
            hCell.SetCellValue(sum);
            
            sum = db.DETECT.Sum(x => x.available);
            hCell = hRow.CreateCell(5);
            hCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
            hCell.SetCellValue(sum);
            
            sum = db.DETECT.Sum(x => x.submit);
            hCell = hRow.CreateCell(6);
            hCell.CellStyle = GetCellStyle(workbook,stylexls.默认);
            hCell.SetCellValue(sum);

            sum = db.DETECT.Sum(x => x.Correct);
            hCell = hRow.CreateCell(7);
            hCell.CellStyle = GetCellStyle(workbook, stylexls.默认);
            hCell.SetCellValue(sum);

            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "各个区域提交汇报.xls");
        }



        /*检查
         * 过程：首先检查总表中的错误信息；检查提交表格 将总表中错误但是提交表中正确的内容更新到总表 假如提交表格中依然还有还有错误，那么在本地目录下保存提交错误表
         */
        public ActionResult Check(string region, string SubmitFile)
        {
            string errorinfomation = null;
            Detect Area = db.DETECT.Where(x => x.region == region).FirstOrDefault();
            if (Area == null)
            {
                return Redirect("/Check/Index");
            }
            /*该用户的数据总表是在Uploads/region目录下的summary.xls*/

            string summaryFile = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region), "summary.xlsx");
            if (Area.submit == 1)
            {
                string message=null;
                Copy(SubmitFile, summaryFile, ref message);

                if (!CollectData(region, ref errorinfomation))
                {
                    Session["Error"] = errorinfomation;
                    return RedirectToAction("Region", new { regionName = region });
                }
            }
            string summaryErrorExcel = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + Area.submit),"summaryErrorExcel.xlsx");
            string subErrorExcel = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + Area.submit), "subErrorExcel.xlsx");
            string AfterExcel = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + Area.submit), "Status.xlsx");
            string Error;
            DetectEngine Engine = new DetectEngine(summaryFile);
            if (!Engine.Check(summaryFile, SubmitFile,AfterExcel, summaryErrorExcel, subErrorExcel,out Error))
            {
                Session["Error"] = "检索失败，失败原因："+Error;
                return RedirectToAction("Region", new { regionName = region });
            }
            if (Engine.subError.Count() != 0)
            {
                ViewBag.sign = "sub";
                ViewBag.name = region;
                return View(Engine.subError);
            }
            if (Engine.summaryError.Count() != 0)
            {
                ViewBag.sign = "summary";
                ViewBag.name = region;
                return View(Engine.summaryError);
            }
            else {
                Area.flag = true;
                db.Entry(Area).State = EntityState.Modified;
                db.SaveChanges();
                return View("success");
            }
        }


        public ActionResult LogOut()
        {
            Session["id"] = null;
            Session["name"] = null;
            return RedirectToAction("Index");
        }


    }
}
