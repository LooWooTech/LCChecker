using LCChecker.Models;
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
            string Catalogue = HttpContext.Server.MapPath("../Uploads/" + CurrentUser.name);
            if (!Directory.Exists(Catalogue))
            {
                try
                {
                    Directory.CreateDirectory(Catalogue);
                }
                catch 
                {
                    throw new ArgumentException("创建目录失败");
                }
            }
            string filePath = null; 
            if (ext == ".xls")
            {
                filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + CurrentUser.name ), "Summary.xls");
            }
            else
            {
                filePath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + CurrentUser.name ), "Summary.xlsx");
            }
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Create);
                fs.Close();
            }
            catch (Exception er)
            {
                throw new ArgumentException(er.Message);
            }
            file.SaveAs(filePath);
            return View();
        }
    }
}
