using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using System;
using System.Collections.Generic;
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

    }
}
