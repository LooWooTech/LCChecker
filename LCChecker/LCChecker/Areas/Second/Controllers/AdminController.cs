using LCChecker.Areas.Second.Models;
using LCChecker.Helpers;
using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Areas.Second.Controllers
{
    public class AdminController : SecondController
    {
        //
        // GET: /Second/Admin/

        public ActionResult Index(City?city,NullableFilter result=NullableFilter.All,int page=1)
        {
            var Filter = new SecondProjectFilter
            {
                City=city,
                Result=result,
                Page=new Page(page)
            };
            ViewBag.List = SecondProjectHelper.GetProjects(Filter);
            ViewBag.Page = Filter.Page;
            //ViewBag.Summary = db.SecondProjects.ToList();
            return View();
        }



        [HttpPost]
        public ActionResult UploadProjects() {
            var file = UploadHelper.GetPostedFile(HttpContext);
            var list = new List<SecondProject>();
            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(1);
            int StartRow = 2;
            //int StartCell = 0;
            int CellIndex=0;
            int Max=sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++) {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                if (string.IsNullOrEmpty(row.Cells[CellIndex+1].ToString())) {
                    continue;
                }

                var id = row.Cells[CellIndex + 4].GetValue().Trim();
                double area = .0;
                double newarea = .0;
                double.TryParse(row.Cells[CellIndex + 12].GetValue().ToString(), out area);
                double.TryParse(row.Cells[CellIndex+20].GetValue().ToString(),out newarea);


                City city = 0;
                var address=row.Cells[CellIndex+3].GetValue().ToString().Replace(',','.').Split('.');
                if (Enum.TryParse<City>(address[1], out city)) {
                    list.Add(new SecondProject
                    {
                        ID=id,
                        City=city,
                        Name=row.Cells[CellIndex+1].GetValue(),
                        County=address[2],
                        Area=area,
                        NewArea=newarea

                    });
                }
            }
            SecondProjectHelper.AddSecondProjects(list);
            return RedirectToAction("Index");
        }


        public ActionResult Statistics() {
            ViewBag.ReportSummary = db.SecondReports.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            ViewBag.CoordSummary = db.CoordProjects.Where(e => e.Visible == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            return View();
        }

        /// <summary>
        /// 下载统计表格
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public ActionResult Download(int ID) {

            Dictionary<City, Summary> Data;
            string ExcelName = "";
            string HeaderName = "";
            if (ID == 1) {
                ExcelName = "报部表格统计.xls";
                HeaderName = "浙江土地整治项目核查报部表格情况统计表";
                Data = db.SecondReports.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
                {
                    TotalCount = g.Count(),
                    ErrorCount = g.Count(e => e.Result == false),
                    SuccessCount = g.Count(e => e.Result == true),
                    City = g.Key
                });
            }
            else if (ID == 2)
            {
                ExcelName = "坐标点存疑统计.xls";
                HeaderName = "浙江省土地整治项目坐标点存疑情况统计表";
                Data = db.CoordProjects.Where(e => e.Visible == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
                {
                    TotalCount = g.Count(),
                    ErrorCount = g.Count(e => e.Result == false),
                    SuccessCount = g.Count(e => e.Result == true),
                    City = g.Key
                });
            }
            else {
                return View();
            }
            IWorkbook workbook = XslHelper.CreateExcel(Data, HeaderName);
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", ExcelName);
        }


        public ActionResult Reports(City city) {
            ViewBag.List = db.SecondReports.Where(e => e.City == city).ToList();
            return View();
        }

        public ActionResult DownCityReport(City city, SecondReportType Type) {
            var uploadFileType = (int)Type+20;
            var file = db.Files.Where(e => e.City == city && e.Type == (UploadFileType)uploadFileType && e.State == UploadFileProceedState.Proceeded).OrderByDescending(e => e.CreateTime).FirstOrDefault();
            if (file == null) {
                throw new Exception("没有找到符合条件的文件");
            }
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, file.SavePath);
            return File(new FileStream(filePath, FileMode.Open), "application/ms-excel", city.ToString() + "-" + Type.ToString() + ".xls");
            //return View();
        }
    }
}
