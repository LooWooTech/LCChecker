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
            int submits = 0;
            Detect record = db.DETECT.Where(x => x.region == regionName).FirstOrDefault();
            if (record == null)
            {
                submits = 0;
            }
            else {
                submits = record.submit;
            }
            ViewBag.submits = submits;
            ViewBag.name = regionName;
            return View();
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
                ViewBag.ErrorMessage = "你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格";
                return RedirectToAction("Region", new { regionName = name });
            }

            if (file.ContentLength == 0||file.ContentLength>20971520)
            {
                ViewBag.ErrorMessage = "你上传的文件0字节或者大于20M 无法读取表格";
                return View();
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
                        ViewBag.ErrorMessage = "服务器本地创建目录失败,错误信息："+er.Message;
                    }  
                }
                
                try
                {
                    FileStream fs = new FileStream(filePath, FileMode.Create);
                    fs.Close();
                }
                catch {
                    ViewBag.ErrorMessage = "服务器保存上传文件失败";
                    return HttpNotFound();
                } 
                file.SaveAs(filePath);
                return RedirectToAction("Check", "Base", new { region = name, SubmitFile=filePath});
            }
        }


        /*
         * 下载存在错误的表格
         * 根据file的名称来判断是下载提交表格中的错误还是总表依然存在的全部错误
         */
        public ActionResult DownErrorExcel(string file)
        {
            if (Session["name"] == null)
            {
                return HttpNotFound();
            }
            string region = Session["name"].ToString();
            Detect Area = db.DETECT.Where(x => x.region == region).FirstOrDefault();
            if (Area == null)
            {
                return RedirectToAction("Index");
            }
            string DownFilePath ;
            string fileName;
            if (file == "sub")
            {
                fileName = "提交错误表.xlsx";
                DownFilePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + Area.submit), "SubError.xlsx");
            }
            else {
                fileName = "总错误表.xlsx";
                DownFilePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + Area.submit), "summaryError.xlsx");
            }
            byte[] fileContents;
            try
            {
                FileStream fs = new FileStream(DownFilePath, FileMode.Open, FileAccess.ReadWrite);
                fileContents = new byte[(int)fs.Length];
                fs.Read(fileContents, 0, fileContents.Length);
                fs.Close();
            }
            catch {
                return HttpNotFound();
            }
            
            
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
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("序号");
            row.CreateCell(1).SetCellValue("区域");
            row.CreateCell(2).SetCellValue("项目个数");
            row.CreateCell(3).SetCellValue("总规模");
            row.CreateCell(4).SetCellValue("新增耕地面积");
            row.CreateCell(5).SetCellValue("可用于占补平衡面积");
            List<Detect> Detects = db.DETECT.ToList();
            int i=1;
            foreach (var item in Detects)
            {
                IRow newRow = sheet.CreateRow(i++);
                newRow.CreateCell(0).SetCellValue(item.Id);
                newRow.CreateCell(1).SetCellValue(item.region);
                newRow.CreateCell(2).SetCellValue(item.sum);
                newRow.CreateCell(3).SetCellValue(item.totalScale);
                newRow.CreateCell(4).SetCellValue(item.AddArea);
                newRow.CreateCell(5).SetCellValue(item.available);
            }
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "区域提交汇报.xls");
        }


    }
}
