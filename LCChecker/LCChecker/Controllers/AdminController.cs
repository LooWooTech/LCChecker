using LCChecker.Models;
using NPOI.HSSF.UserModel;
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
        private void AddProjects(List<Project> list)
        {
            foreach (var item in list)
            {
                if (!db.Projects.Any(e => e.ID == item.ID))
                {
                    db.Projects.Add(item);
                }
                db.SaveChanges();
            }
        }

        public ActionResult Index()
        {
            return View();
        }



        /// <summary>
        /// 上传总表数据
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects()
        {
            if (Request.Files.Count == 0)
            {
                throw new ArgumentException("请选择文件上传");
            }

            HttpPostedFileBase file = Request.Files[0];
            string ext = Path.GetExtension(file.FileName);
            if (ext != ".xls" || ext != "xlsx")
            {
                throw new ArgumentException("你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格");
            }
            if (file.ContentLength == 0 || file.ContentLength > 20971520)
            {
                throw new ArgumentException("你上传的文件数据太大或者没有");
            }

            var list = new List<Project>();

            var excel = new HSSFWorkbook(file.InputStream);
            var sheet = excel.GetSheetAt(0);
            for (var i = 1; i < sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (string.IsNullOrEmpty(row.Cells[0].StringCellValue))
                {
                    continue;
                }
                var cityNames = row.Cells[0].StringCellValue.Split(',');
                if (cityNames.Length < 2)
                {
                    continue;
                }

                City city = 0;

                if (Enum.TryParse<City>(cityNames[1], out city))
                {
                    list.Add(new Project
                    {
                        ID = row.Cells[1].StringCellValue,
                        City = city
                    });
                }

            }

            AddProjects(list);

            return RedirectToAction("Index");
        }
    }
}
