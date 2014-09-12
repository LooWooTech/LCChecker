using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker
{
    public class UploadHelper
    {
        private static string UploadDirectory = "Uploads/";

        private static string GetAbsoluteUploadDirectory(string fileName)
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, UploadDirectory, fileName);
        }

        public static string GetAbsolutePath(string filePath)
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePath);
        }

        public static HttpPostedFileBase GetPostedFile(HttpContextBase context)
        {
            if (context.Request.Files.Count == 0)
            {
                throw new ArgumentException("请选择文件上传");
            }

            HttpPostedFileBase file = null;
            for (var i = 0; i < context.Request.Files.Count; i++)
            {
                file = context.Request.Files[i];
                if (file.ContentLength > 0)
                {
                    break;
                }
            }
                //foreach (HttpPostedFileBase file1 in context.Request.Files)
                //{
                //    if (file1.ContentLength > 0)
                //    {
                //        file = file1;
                //        break;
                //    }
                //}
            return file;
        }


        public static string UploadExcel(HttpPostedFileBase file)
        {

            var ext = Path.GetExtension(file.FileName);
            if (ext != ".xls" && ext != ".xlsx")
            {
                throw new ArgumentException("你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格");
            }
            if (file.ContentLength == 0 || file.ContentLength > 20971520)
            {
                throw new ArgumentException("你上传的文件数据太大或者没有");
            }

            var fileName =  DateTime.Now.Ticks.ToString() + ext;

            file.SaveAs(GetAbsoluteUploadDirectory(fileName));

            return UploadDirectory + fileName;
        }
    }
}