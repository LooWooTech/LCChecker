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



        public ActionResult Admin()
        {

            return View();
        }


        public ActionResult Region(string regionName)
        {
            ViewBag.name = regionName;
            return View();
        }


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
                FileStream fs = new FileStream(filePath, FileMode.Create);
                fs.Close();
                file.SaveAs(filePath);
                return RedirectToAction("shangchuan", new { fPath = filePath });

            }
            return RedirectToAction("Region", new { regionName = name });
        }


        public ActionResult shangchuan(string fPath)
        {
            List<Mistake> information = CheckExcel(fPath);
            return View(information);
        }

        //public FileResult Download()
        //{
        //    MemoryStream ms = new MemoryStream();
        //    IWorkbook workbook = new HSSFWorkbook();
        //    ISheet sheet = workbook.CreateSheet();
        //    IRow Row1 = sheet.CreateRow(0);
        //}



       
    }
}
