using LCChecker.Helpers;
using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    [UserAuthorize]
    public class AdminController : ControllerBase
    {
        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (!CurrentUser.Flag)
            {
                throw new HttpException(401, "你没有权限查看此页面");
            }
            base.OnActionExecuting(filterContext);
        }

        public ActionResult Index()
        {
            ViewBag.Summary = db.Projects.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            return View();
        }

        public ActionResult Users()
        {
            ViewBag.List = db.Users.ToList();
            return View();
        }

        [HttpGet]
        public ActionResult AddUser()
        {
            return View();
        }
        [HttpPost]
        public ActionResult AddUser(User user)
        {
            if (string.IsNullOrEmpty(user.Username))
            {
                return JsonError("用户名没有填写");
            }

            if (string.IsNullOrEmpty(user.Password))
            {
                return JsonError("密码没有填写");
            }

            var rePwd = Request.Form["repassword"];
            if (string.IsNullOrEmpty(rePwd))
            {
                return JsonError("确认密码没有填写");
            }

            if (rePwd != user.Password)
            {
                return JsonError("密码输入不一致");
            }

            if (user.City == 0 && user.Flag == false)
            {
                return JsonError("非管理员用户请选择城市");
            }

            if (db.Users.Any(e => e.Username.ToLower() == user.Username.ToLower()))
            {
                return JsonError("用户名已被占用");
            }

            db.Users.Add(user);
            db.SaveChanges();

            return JsonSuccess();
        }

        public ActionResult DeleteUser(int id)
        {
            try
            {
                var entity = db.Users.FirstOrDefault(e => e.ID == id);
                db.Users.Remove(entity);
                db.SaveChanges();
                return JsonSuccess();
            }
            catch (Exception ex)
            {
                return JsonError(ex.Message);
            }
        }

        public ActionResult Projects(City? city, NullableFilter result = NullableFilter.All, int page = 1)
        {
            var filter = new ProjectFileter
            {
                City = city,
                Result = result,
                Page = new Page(page)
            };
            ViewBag.List = ProjectHelper.GetProjects(filter);
            ViewBag.Page = filter.Page;
            return View();
        }

        public ActionResult Reports(City city)
        {
            ViewBag.List = db.Reports.Where(e => e.City == city).ToList();
            return View();
        }

        public ActionResult DownloadReport(City city, ReportType type)
        {
            var uploadFileType = (int)type;
            var file = db.Files.Where(e => e.City == city && e.Type == (UploadFileType)uploadFileType && e.State == UploadFileProceedState.Proceeded).OrderByDescending(e => e.CreateTime).FirstOrDefault();
            if (file == null)
            {
                throw new Exception("没有找到符合条件的文件");
            }
            return File(Request.MapPath(file.SavePath), "application/ms-excel", city.ToString() + "-" + type.ToString() + ".xls");
        }

        [HttpGet]
        public ActionResult Statistics()
        {
            ViewBag.ProjectSummary = db.Projects.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });

            ViewBag.ReportSummary = db.Reports.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
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

        public ActionResult DownLoadExcel(int id)
        {
            string[] Header = { "行政区", "项目总数", "通过总数", "失败总数", "未上传数" };
            int[] He = new int[4];
            Dictionary<City, Summary> Data;
            string ExcelName = "";
            string HeaderName = "";
            if (id == 1)
            {
                ExcelName = "自检表统计表格.xls";
                HeaderName = "浙江省土地整治项目核查自检情况统计表";
                Data = db.Projects.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
                {
                    TotalCount = g.Count(),
                    ErrorCount = g.Count(e => e.Result == false),
                    SuccessCount = g.Count(e => e.Result == true),
                    City = g.Key
                });
            }
            else if (id == 2)
            {
                ExcelName = "报部表格统计.xls";
                HeaderName = "浙江省土地整治项目核查报部情况统计表";
                Data = db.Reports.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
                {
                    TotalCount = g.Count(),
                    ErrorCount = g.Count(e => e.Result == false),
                    SuccessCount = g.Count(e => e.Result == true),
                    City = g.Key
                });
            }
            else if (id == 3)
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
            else
            {
                ExcelName = "报部表格统计.xls";
                Data = db.Reports.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
                {
                    TotalCount = g.Count(),
                    ErrorCount = g.Count(e => e.Result == false),
                    SuccessCount = g.Count(e => e.Result == true),
                    City = g.Key
                });
            }
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("sheet1");
            for (var j = 0; j < 5; j++)
            {
                sheet.SetColumnWidth(j, 15 * 256);
            }


            if (id == 4)
            {
                Dictionary<string, Dictionary<string, Dictionary<string, int>>> Message = new Dictionary<string, Dictionary<string, Dictionary<string, int>>>();
                summ(ref Message);
                IRow row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("市");
                row.CreateCell(1).SetCellValue("项目总数");
                row.CreateCell(2).SetCellValue("通过总数");
                row.CreateCell(3).SetCellValue("失败总数");
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 3, 5));
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 0));
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 1, 1));
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 2, 2));
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 6, 6));
                row.CreateCell(6).SetCellValue("未上传数");
                row = sheet.CreateRow(1);
                row.CreateCell(3).SetCellValue("总数");
                row.CreateCell(4).SetCellValue("提示");
                row.CreateCell(5).SetCellValue("错误");
                int i = 2;
                foreach (var item in Data)
                {
                    var summary = item.Value;
                    row = sheet.CreateRow(i++);
                    row.CreateCell(0).SetCellValue(summary.City.ToString());
                    row.CreateCell(1).SetCellValue(summary.TotalCount);
                    row.CreateCell(2).SetCellValue(summary.SuccessCount);
                    row.CreateCell(3).SetCellValue(summary.ErrorCount);
                    int all = 0;
                    if (Message.ContainsKey(summary.City.ToString()))
                    {
                        foreach (var it in Message[summary.City.ToString()].Keys)
                        {
                            if (Message[summary.City.ToString()][it].ContainsKey("错误"))
                            {
                                all++;
                            }
                        }

                    }
                    row.CreateCell(4).SetCellValue(summary.ErrorCount - all);
                    row.CreateCell(5).SetCellValue(all);
                    row.CreateCell(6).SetCellValue(summary.UnCheckCount);
                }

            }
            else
            {
                IRow row = sheet.CreateRow(0);
                var cell = row.CreateCell(0);
                cell.CellStyle = workbook.GetCellStyle(XslHeaderStyle.大头);
                cell.SetCellValue(HeaderName);
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 4));
                row = sheet.CreateRow(1);
                int i = 0;
                foreach (var item in Header)
                {
                    cell = row.CreateCell(i++);
                    cell.CellStyle = workbook.GetCellStyle(XslHeaderStyle.小头);
                    cell.SetCellValue(item);
                }
                i = 2;
                foreach (var item in Data)
                {
                    var summary = item.Value;
                    row = sheet.CreateRow(i++);
                    for (var j = 0; j < 5; j++)
                    {
                        cell = row.CreateCell(j);
                        cell.CellStyle = workbook.GetCellStyle(XslHeaderStyle.默认);
                    }
                    row.GetCell(0).SetCellValue(summary.City.ToString());
                    row.GetCell(1).SetCellValue(summary.TotalCount);
                    He[0] += summary.TotalCount;
                    row.GetCell(2).SetCellValue(summary.SuccessCount);
                    He[1] += summary.SuccessCount;
                    row.GetCell(3).SetCellValue(summary.ErrorCount);
                    He[2] += summary.ErrorCount;
                    row.GetCell(4).SetCellValue(summary.UnCheckCount);
                    He[3] += summary.UnCheckCount;
                }
                row = sheet.CreateRow(i++);
                for (var j = 0; j < 5; j++)
                {
                    cell = row.CreateCell(j);
                    cell.CellStyle = workbook.GetCellStyle(XslHeaderStyle.默认);
                    if (j != 0)
                    {
                        cell.SetCellValue(He[j - 1]);
                    }
                }
                row.GetCell(0).SetCellValue("合计");

            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", ExcelName);
        }

        public void summ(ref Dictionary<string, Dictionary<string, Dictionary<string, int>>> Error)
        {
            // Dictionary<string, Dictionary<string, Dictionary<string, int>>> Error = new Dictionary<string, Dictionary<string, Dictionary<string, int>>>();
            foreach (City item in Enum.GetValues(typeof(City)))
            {
                string chengshi = item.ToString();
                Dictionary<string, Dictionary<string, int>> Baobu = new Dictionary<string, Dictionary<string, int>>();
                var Records = db.Records.Where(x => x.City == item).ToList();
                foreach (var cord in Records)
                {
                    if (!string.IsNullOrEmpty(cord.Note))
                    {
                        string Type = cord.Type.ToString();
                        string Message = "";
                        if (cord.IsError == false)
                        {
                            Message = "提示";
                        }
                        else
                        {
                            Message = "错误";
                        }
                        if (Baobu.ContainsKey(Type))
                        {
                            if (Baobu[Type].ContainsKey(Message))
                            {
                                Baobu[Type][Message]++;
                            }
                            else
                            {
                                Baobu[Type].Add(Message, 1);
                            }
                        }
                        else
                        {
                            Dictionary<string, int> Fault = new Dictionary<string, int>();
                            Fault.Add(Message, 1);
                            Baobu.Add(Type, Fault);
                        }
                    }
                }
                if (!Error.ContainsKey(chengshi) && Baobu.Count != 0)
                {
                    Error.Add(chengshi, Baobu);
                }
            }
        }

        /// <summary>
        /// 上传总表数据（导入表3）
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects()
        {
            var file = UploadHelper.GetPostedFile(HttpContext);

            var list = new List<Project>();

            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(0);

            int rowIndex = 0, cellIndex = 0;
            var title = sheet.GetRow(rowIndex).GetCell(cellIndex).GetValue();
            if (title != "附表3")
            {
                throw new ArgumentException("上传的附表3：重点项目复核确认总表格式不正确，请参照样表。");
            }
            rowIndex = rowIndex + 6;
            for (var i = rowIndex; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                if (string.IsNullOrEmpty(row.Cells[cellIndex].ToString()))
                {
                    continue;
                }
                var id = row.Cells[cellIndex + 3].GetValue().Trim();
                //if (!id.VerificationID())
                //{
                //    continue;
                //}

                City city = 0;

                if (Enum.TryParse<City>(row.Cells[cellIndex + 1].GetValue(), out city))
                {
                    list.Add(new Project
                    {
                        City = city,
                        County = row.Cells[cellIndex + 2].GetValue(),
                        ID = id,
                        Name = row.Cells[cellIndex + 4].GetValue(),
                        IsHasError = row.Cells[cellIndex + 5].GetValue() == "否",
                        IsApplyDelete = row.Cells[cellIndex + 6].GetValue() == "是",
                        IsShouldModify = row.Cells[cellIndex + 7].GetValue() == "是",
                        IsDecrease = row.Cells[cellIndex + 9].GetValue() == "是",

                    });
                }

            }

            ProjectHelper.AddProjects(list);

            return RedirectToAction("Projects", "Admin");
        }

        public ActionResult CoordProjects(City? city, NullableFilter result = NullableFilter.All, int page = 1)
        {
            var filter = new ProjectFileter
            {
                City = city,
                Result = result,
                Visible = true,
                Page = new Page(page)
            };
            ViewBag.List = ProjectHelper.GetCoordProjects(filter);
            ViewBag.Page = filter.Page;
            return View();
        }

        [HttpPost]
        public ActionResult UploadCoords()
        {
            var file = UploadHelper.GetPostedFile(HttpContext);


            var excel = XslHelper.GetWorkbook(file);
            var sheet = excel.GetSheetAt(0);

            var list = new List<CoordProject>();

            var visible = false;

            var firstCell = sheet.GetRow(0).GetCell(0);
            if (firstCell.GetValue() == "市")
            {
                visible = true;
            }

            var rowIndex = visible ? 1 : 7;
            var cellIndex = visible ? 0 : 1;
            for (; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (string.IsNullOrEmpty(row.Cells[0].ToString()))
                {
                    continue;
                }

                if (!visible && row.Cells[6].StringCellValue == "是")
                {
                    continue;
                }

                City city = 0;

                if (Enum.TryParse<City>(row.Cells[cellIndex].GetValue(), out city))
                {
                    list.Add(new CoordProject
                    {
                        City = city,
                        County = row.Cells[cellIndex + 1].GetValue(),
                        ID = row.Cells[cellIndex + 2].GetValue(),
                        Name = row.Cells[cellIndex + 3].GetValue(),
                        Note = visible ? row.Cells[cellIndex + 4].GetValue() : null,
                        Visible = visible
                    });
                }
            }
            ProjectHelper.AddCoordProjects(list);
            return RedirectToAction("CoordProjects");
        }
    }
}