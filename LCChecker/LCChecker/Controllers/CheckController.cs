using LCChecker.Models;
using LCChecker.Rules;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml;

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
                filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + name + "/" + record.submit), "NO" + record.submit + ".xlsx");
                string Catalogue = HttpContext.Server.MapPath("../Uploads/" + name+"/"+record.submit);
                Directory.CreateDirectory(Catalogue);
                HttpPostedFileBase files = Request.Files[0];
                using (FileStream fs = new FileStream(filePath, FileMode.Create)) { }
                file.SaveAs(filePath);
                return RedirectToAction("Check", "Base", new { region = name });
            }
        }


        //public ActionResult jiancha(string fPath)
        //{
        //    if(Session["name"]==null)
        //        return HttpNotFound();
            
        //    string name=Session["name"].ToString();
        //    Detect record=db.DETECT.Where(x=>x.region==name).FirstOrDefault();
        //    if(record==null)
        //    {
        //        return HttpNotFound();
        //    }

        //    string xmlPath=Path.Combine(HttpContext.Server.MapPath("../Uploads/"+name),record.submit+".xml");
        //    List<Mistake> information = CheckExcel(@"E:\LCChecker\trunk\LCChecker\LCChecker\Uploads\湖州市\No8.xls",name,record.submit);
        //    XmlWriterSettings settings = new XmlWriterSettings();
        //    settings.Indent = true;
        //    settings.NewLineOnAttributes = true;
        //    XmlWriter writer = XmlWriter.Create(xmlPath, settings);
        //    writer.WriteStartDocument();
        //    foreach (var item in information)
        //    {
        //        //if (item.flag)
        //        //{
        //        //    writer.WriteStartElement("Error");
        //        //    writer.WriteElementString("ErrorType", item.Error);
        //        //    writer.WriteElementString("row", item.row.ToString());
        //        //    writer.WriteEndElement();
        //        //}
        //    }
        //    writer.WriteEndDocument();
            
        //    return View(information);
        //}

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

        /*区域用户下载错误表格*/
        //public FileResult MistakeDown(string name)
        //{
        //    Detect record = db.DETECT.Where(x => x.region == name).FirstOrDefault();
        //    string xmlPath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + name), record.submit + ".xml");
        //    string filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + name), name + ".xls");
        //    List<Mistake> mistakes = new List<Mistake>();
        //    XmlReaderSettings settings = new XmlReaderSettings();
        //    XmlReader rdr = XmlReader.Create(xmlPath);
        //    while (rdr.Read())
        //    {
        //        if (rdr.NodeType == XmlNodeType.Text)
        //        {
        //            mistakes.Add(new Mistake() { Error = rdr.Value, row = int.Parse(rdr.Value) });
        //        }
        //    }
        //    int count = mistakes.Count();
        //    MemoryStream ms = new MemoryStream();
        //    IWorkbook workbook = new HSSFWorkbook();
        //    ISheet sheet = workbook.CreateSheet();
        //    IRow row = sheet.CreateRow(0);
        //    for (int i = 0; i < 43; i++)
        //    {
        //        row.CreateCell(i).SetCellValue(string.Format("{0}栏", i + 1));
        //    }
            
        //}


    }
}
