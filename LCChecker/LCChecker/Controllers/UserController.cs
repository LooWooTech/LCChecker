using LCChecker.Models;
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
            else {
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
            if (summary.TotalCount == summary.SuccessCount)
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
            string CheckSelfExcel = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates/", "ModelSelf.xlsx");
            IWorkbook workbook;
            try
            {
                FileStream fs = new FileStream(CheckSelfExcel, FileMode.Open, FileAccess.Read);
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
                cell2.SetCellValue(City.浙江省.ToString() + "," + item.City.ToString());

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

            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,file.SavePath);
            //读取文件进行检查
            Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();
            Dictionary<string, int> ship = new Dictionary<string, int>();
            DetectEngine Engine = new DetectEngine(filePath);
            string fault="";
            if (!Engine.CheckExcel(filePath, ref fault, ref Error, ref ship))
            {
                throw new ArgumentException("检索失败");
            }
            //检查完毕，更新Projects
            var projects = db.Projects.Where(x => x.City == CurrentUser.City).ToList();
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
                    else {
                        item.Result = true;
                        item.Note = "";
                    }
                    db.Entry(item).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }

            return RedirectToAction("Index", new { result=false});
        }
    }
}
