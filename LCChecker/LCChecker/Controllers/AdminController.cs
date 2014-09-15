﻿using LCChecker.Models;
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
            ViewBag.Summary = db.Projects.GroupBy(e => e.City).ToDictionary(g => g.Key, g => new Summary
            {
                TotalCount = g.Count(),
                ErrorCount = g.Count(e => e.Result == false),
                SuccessCount = g.Count(e => e.Result == true),
                City = g.Key
            });
            return View();
        }

        public ActionResult Projects(City? city, int page = 1)
        {
            var paging = new Page(page);
            var query = db.Projects.AsQueryable();
            if (city.HasValue)
            {
                query = query.Where(e => e.City == city.Value);
            }
            paging.RecordCount = query.Count();
            ViewBag.List = query.OrderBy(e => e.ID).Skip(paging.PageSize * (paging.PageIndex - 1)).Take(paging.PageSize).ToList();
            ViewBag.Page = paging;
            return View();
        }

        /// <summary>
        /// 上传总表数据
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects()
        {
            var file = UploadHelper.GetPostedFile(HttpContext);

            var list = new List<Project>();

            IWorkbook excel = null;
            var ext = Path.GetExtension(file.FileName);
            if (ext == ".xls")
            {
                excel = new HSSFWorkbook(file.InputStream);
            }
            else
            {
                excel = new XSSFWorkbook(file.InputStream);
            }
            //var excel = new HSSFWorkbook(file.InputStream);
            var sheet = excel.GetSheetAt(0);
            for (var i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (string.IsNullOrEmpty(row.Cells[0].StringCellValue))
                {
                    continue;
                }
                var cityNames = row.Cells[0].StringCellValue.Split(',');
                if (cityNames.Length < 3)
                {
                    continue;
                }

                var county = cityNames[2];

                City city = 0;

                if (Enum.TryParse<City>(cityNames[1], out city))
                {
                    list.Add(new Project
                    {
                        City = city,
                        ID = row.Cells[1].NumericCellValue.ToString(),
                        Name = row.Cells[2].StringCellValue,
                        County = county
                    });
                }

            }

            AddProjects(list);

            return RedirectToAction("Index", "Admin");
        }
    }
}