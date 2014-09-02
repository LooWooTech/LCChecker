using LCChecker.Models;
using LCChecker.Rules;
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

        public int beginCell = 0;
        public Dictionary<string, int> Relatship = new Dictionary<string, int>(43);
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

            return View();
        }

        /*区域用户登录*/
        public ActionResult Region(string regionName)
        {
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
                return View();
            }

            var file = Request.Files[0];
            if (file.ContentLength == 0)
            {
                return View();
            }
            else {
                HttpPostedFileBase files = Request.Files[0];
                string Catalogue = HttpContext.Server.MapPath("../Uploads/"+name);
                if (!Directory.Exists(Catalogue))
                {
                    Directory.CreateDirectory(Catalogue);
                }
                Detect record = db.DETECT.Where(x => x.region == name).FirstOrDefault();
                string filePath = null;
                if (record == null)
                {
                    Detect NewRecord = new Detect();
                    NewRecord.region = name;
                    NewRecord.submit = 1;
                    if (ModelState.IsValid)
                    {
                        db.DETECT.Add(NewRecord);
                        db.SaveChanges();
                    }
                    filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + name), name+".xls");
                }
                else {
                    record.submit++;
                    if (ModelState.IsValid)
                    {
                        db.Entry(record).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + name), "No" + record.submit + ".xls");

                }
                using (FileStream fs = new FileStream(filePath, FileMode.Create)) { }
                file.SaveAs(filePath);
                return RedirectToAction("shangchuan", new { fPath = filePath });

            }
            return RedirectToAction("Region", new { regionName = name });
        }


        public ActionResult jiancha(string fPath)
        {
            List<Mistake> information = CheckExcel(@"E:\LCChecker\trunk\LCChecker\LCChecker\Uploads\湖州市\No8.xls");
            return View(information);
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
