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

        public ActionResult PlanIndex(City? city, NullableFilter result = NullableFilter.All, int page = 1) {
            var filter = new SecondProjectFilter
            {
                City = city,
                Result = result,
                Page = new Page(page)
            };
            ViewBag.List = SecondProjectHelper.GetPlanProjects(filter);
            ViewBag.Page = filter.Page;
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
                double SurplusHookArea = .0;
                double TrueHookArea = 0.0;
                double.TryParse(row.Cells[CellIndex + 12].GetValue().ToString(), out area);
                double.TryParse(row.Cells[CellIndex+20].GetValue().ToString(),out newarea);
                double.TryParse(row.Cells[CellIndex + 5].GetValue().ToString(), out SurplusHookArea);
                double.TryParse(row.Cells[CellIndex + 21].GetValue().ToString(), out TrueHookArea);

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
                        NewArea=newarea,
                        SurplusHookArea=SurplusHookArea,
                        TrueHookArea=TrueHookArea
                    });
                }
            }
            SecondProjectHelper.AddSecondProjects(list);
            return RedirectToAction("Index");
        }
        [HttpPost]
        public ActionResult UploadPlanProject() {
            var file = UploadHelper.GetPostedFile(HttpContext);
           // var list = new List<SecondProject>();
            var plist = new List<pProject>();
            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(0);
            int StartRow = 1;
            int Max = sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++) {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                var value = row.GetCell(7).ToString().Trim().ToUpper(); ;
                if (string.IsNullOrEmpty(value))
                    continue;
                //if (!SecondProjectHelper.VerificationPID(value))
                //    continue;
                City city = 0;
                var address = row.Cells[3].GetValue().ToString().Replace('，', '.').Split('.');
                if (Enum.TryParse<City>(address[1], out city)) {
                    plist.Add(new pProject
                    {
                        ID = value,
                        City = city,
                        Name = row.Cells[2].GetValue(),
                        County = address[2],
                    });
                }
            }
            //SecondProjectHelper.AddSecondProjects(list);
            SecondProjectHelper.AddPlanProject(plist);

            return RedirectToAction("PlanIndex");
        }



        public ActionResult NewAreaCoord(City?city,NullableFilter result=NullableFilter.All,int page=1) {
            var Filter = new SecondProjectFilter
            {
                City = city,
                Result = result,
                Page = new Page(page)
            };
            ViewBag.List = SecondProjectHelper.GetNewAreaCoord(Filter);
            ViewBag.Page = Filter.Page;
            return View();
        }
        [HttpPost]
        public ActionResult UploadNewAreaCoord()
        {
            var file = UploadHelper.GetPostedFile(HttpContext);
            var list = new List<CoordNewAreaProject>();
            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(1);
            int StartRow = 2;
            //int StartCell = 0;
            int CellIndex = 0;
            int Max = sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                if (string.IsNullOrEmpty(row.Cells[CellIndex + 1].ToString()))
                {
                    continue;
                }

                var id = row.Cells[CellIndex + 4].GetValue().Trim();


                City city = 0;
                var address = row.Cells[CellIndex + 3].GetValue().ToString().Replace(',', '.').Split('.');
                if (Enum.TryParse<City>(address[1], out city))
                {
                    list.Add(new CoordNewAreaProject
                    {
                        ID = id,
                        City = city,
                        Name=row.Cells[CellIndex+1].GetValue(),
                        County = address[2],
                        Visible = true
                    });
                }
            }
            SecondProjectHelper.AddCoordNewArea(list);
            return RedirectToAction("NewAreaCoord");
        }
        [HttpPost]
        public ActionResult UploadFirstNewArea() {
            var file = UploadHelper.GetPostedFile(HttpContext);
            var list = new List<CoordNewAreaProject>();
            var excel = XslHelper.GetWorkbook(file);
            var sheet1 = excel.GetSheetAt(0);
            int Max = sheet1.LastRowNum;
            int StartRow = 1;
            for (var i = StartRow; i <= Max; i++) {
                var row = sheet1.GetRow(i);
                if (row == null)
                    continue;
                var value = row.Cells[3].GetValue().ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                if (!value.VerificationID())
                    continue;
                City city = 0;
                var address = row.Cells[1].GetValue().ToString().Trim();
                if (Enum.TryParse<City>(address, out city)) {
                    list.Add(new CoordNewAreaProject
                    {
                        ID = value,
                        City = city,
                        Name = row.Cells[4].GetValue().ToString().Trim(),
                        County = row.Cells[2].GetValue().ToString().Trim(),
                        Visible = true
                    });
                }
                
            }
            SecondProjectHelper.AddCoordNewArea(list);
            return RedirectToAction("NewAreaCoord");
        }



        [HttpPost]
        public ActionResult UploadCoord() {
            var file = UploadHelper.GetPostedFile(HttpContext);
            var list = new List<CoordProject>();
            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(1);
            int StartRow = 2;
            //int StartCell = 0;
            int CellIndex = 0;
            int Max = sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                if (string.IsNullOrEmpty(row.Cells[CellIndex + 1].ToString()))
                {
                    continue;
                }

                var id = row.Cells[CellIndex + 4].GetValue().Trim();


                City city = 0;
                var address = row.Cells[CellIndex + 3].GetValue().ToString().Replace(',', '.').Split('.');
                if (Enum.TryParse<City>(address[1], out city))
                {
                    list.Add(new CoordProject
                    {
                        ID = id,
                        City = city,
                        Name = row.Cells[CellIndex + 1].GetValue(),
                        County = address[2],
                        Visible = false,
                    });
                }
            }
            SecondProjectHelper.AddCoordProject(list);
            return Redirect("/Admin/CoordProjects");
            //return RedirectToAction("CoordProjects","Admin");
        }


        public ActionResult Statistics() {
            ViewBag.ReportSummary = db.SecondReports.Where(e=>e.IsPlan==false).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            ViewBag.PlanReportSummary = db.SecondReports.Where(e => e.IsPlan == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            ViewBag.CoordSummary = db.CoordProjects.Where(e => e.Visible == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false&&e.Exception==false),
                ExceptionCount=g.Count(e=>e.Exception==true),
                SuccessCount = g.Count(e => e.Result == true&&e.Exception==false),
                City = g.Key
            });
            ViewBag.CoordNewAreaSummary = db.CoordNewAreaProjects.Where(e => e.Visible == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false && e.Exception == false),
                ExceptionCount = g.Count(e => e.Exception == true),
                SuccessCount = g.Count(e => e.Result == true&&e.Exception==false),
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
            bool flag = false;
            if (ID == 1) {
                flag = true;
                ExcelName = "报部表格统计.xls";
                HeaderName = "浙江土地整治项目核查验收报部表格情况统计表";
                Data = db.SecondReports.Where(e=>e.IsPlan==false).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
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
                    ErrorCount = g.Count(e => e.Result == false&&e.Exception==false),
                    ExceptionCount=g.Count(e=>e.Exception==true),
                    SuccessCount = g.Count(e => e.Result == true&&e.Exception==false),
                    City = g.Key
                });
            }
            else if (ID == 3) {
                ExcelName = "新增耕地坐标统计.xls";
                HeaderName = "浙江省土地整治项目新增耕地坐标情况统计表";
                Data = db.CoordNewAreaProjects.Where(e => e.Visible == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
                {
                    TotalCount = g.Count(),
                    ErrorCount = g.Count(e => e.Result == false && e.Exception == false),
                    ExceptionCount = g.Count(e =>e.Exception == true),
                    SuccessCount = g.Count(e => e.Result == true&&e.Exception==false),
                    City = g.Key
                });
            }
            else if (ID == 4) {
                flag = true;
                ExcelName = "报部表格统计.xls";
                HeaderName = "浙江土地整治项目核查未验收报部表格情况统计表";
                Data = db.SecondReports.Where(e => e.IsPlan == true).GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
                {
                    TotalCount = g.Count(),
                    ErrorCount = g.Count(e => e.Result == false),
                    SuccessCount = g.Count(e => e.Result == true),
                    City = g.Key
                });
            }
            else
            {
                return View();
            }
            IWorkbook workbook = XslHelper.CreateExcel(Data, HeaderName,flag);
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", ExcelName);
        }


        public ActionResult Reports(City city,bool IsPlan) {
            ViewBag.List = db.SecondReports.Where(e => e.City == city&&e.IsPlan==IsPlan).ToList();
            return View();
        }

        public ActionResult DownCityReport(City city, SecondReportType Type,bool IsPlan) {
            var uploadFileType = (int)Type+20;
            var file = db.Files.Where(e => e.City == city && e.Type == (UploadFileType)uploadFileType && e.State == UploadFileProceedState.Proceeded&&e.IsPlan==IsPlan).OrderByDescending(e => e.CreateTime).FirstOrDefault();
            if (file == null) {
                throw new Exception("没有找到符合条件的文件");
            }
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, file.SavePath);
            return File(new FileStream(filePath, FileMode.Open), "application/ms-excel", city.ToString() + "-" + Type.ToString() + ".xls");
            //return View();
        }
    }
}
