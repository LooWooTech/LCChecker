﻿using ICSharpCode.SharpZipLib.Zip;
using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
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
    public partial class UserController : ControllerBase
    {
        public ActionResult Index()
        {

            ViewBag.ProjectSummary = new Summary
            {
                City = CurrentUser.City,
                TotalCount = db.Projects.Count(e => e.City == CurrentUser.City),
                SuccessCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == true),
                ErrorCount = db.Projects.Count(e => e.City == CurrentUser.City && e.Result == false),

            };
            ViewBag.ReportSummary = new Summary
            {
                City = CurrentUser.City,
                TotalCount = db.Reports.Count(e => e.City == CurrentUser.City),
                SuccessCount = db.Reports.Count(e => e.City == CurrentUser.City && e.Result == true),
                ErrorCount = db.Reports.Count(e => e.City == CurrentUser.City && e.Result == false),
            };
            return View();
        }

        public ActionResult Projects(NullableFilter result = NullableFilter.All, int page = 1)
        {
            var filter = new ProjectFileter
            {
                City = CurrentUser.City,
                Result = result,
                Page = new Page(page),
            };
            ViewBag.List = ProjectHelper.GetProjects(filter);
            ViewBag.Page = filter.Page;
            //ViewBag.ProjectType = filter.Type;
            return View();
        }

        /// <summary>
        /// 上传一部分项目，验证并更新到Project
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UploadProjects(/*ProjectType type*/)
        {
            //if (!CurrentUser.Flag)
            //{
            //    throw new ArgumentException("已经超出上传时间");
            //}
            
            var file = UploadHelper.GetPostedFile(HttpContext);

            var ext = Path.GetExtension(file.FileName);
            if (ext != ".xls" && ext != ".xlsx")
            {
                throw new ArgumentException("你上传的文件格式不对，目前支持.xls以及.xlsx格式的EXCEL表格");
            }

            var filePath = UploadHelper.Upload(file);

            var uploadFile = new UploadFile
            {
                City = CurrentUser.City,
                CreateTime = DateTime.Now,
                FileName = file.FileName,
                SavePath = filePath
            };

            var fileId = UploadHelper.AddFileEntity(uploadFile);

            return RedirectToAction("CheckProject", new { id = fileId/*, type = (int)type */});
        }

        public ActionResult CheckProject(int id/*, ProjectType type*/)
        {
            var file = db.Files.FirstOrDefault(e => e.ID == id);
            if (file == null)
            {
                throw new ArgumentException("参数错误");
            }
            else
            {
                file.State = UploadFileProceedState.Proceeded;
            }

            var filePath = UploadHelper.GetAbsolutePath(file.SavePath);
            var MyProjects = db.Projects.Where(e => e.City == CurrentUser.City).ToList();
            //读取文件进行检查
            var errors = new Dictionary<string, List<string>>();
            var ships = new Dictionary<string, int>();
            var areas=new Dictionary<string,double[]>();
            var detectEngine = new DetectEngine(filePath,MyProjects);
            var fault = "";
            if (!detectEngine.CheckExcel(filePath, ref fault, ref errors, ref ships, ref areas))
            {
                file.State = UploadFileProceedState.Error;
                file.ProcessMessage = "检索失败：" + fault;
                db.SaveChanges();
                throw new ArgumentException("检索失败：" + fault);
            }

            var masterfile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/", CurrentUser.City.ToString() + ".xls");

            var list = db.Projects.Where(x => x.City == CurrentUser.City).ToList();

            if (!detectEngine.SaveCurrent(filePath, masterfile, ref fault, errors, ships, list))
            {
                throw new ArgumentException("保存正确项目失败");
            }

            //检查完毕，更新Projects
            var projects = db.Projects.Where(e => e.City == CurrentUser.City);
            var checkTime = DateTime.Now;
            foreach (var item in projects)
            {
                if (ships.ContainsKey(item.ID))
                {
                    if (errors.ContainsKey(item.ID))
                    {
                        item.Note = "";
                        item.Result = false;
                        var errs = errors[item.ID];
                        var i = 1;
                        foreach (var msg in errs)
                        {
                            item.Note += string.Format("({0}){1}；", i, msg); ;
                            i++;
                        }
                    }
                    else
                    {
                        if(areas.ContainsKey(item.ID))//更新总规模  新增耕地面积  在
                        {
                            item.Area=areas[item.ID][0];
                            item.NewArea=areas[item.ID][1];
                        }
                        item.Result = true;
                        item.Note = "";
                    }
                    item.UpdateTime = checkTime;
                }
            }
            if (errors.Count > 0)
            {
                file.State = UploadFileProceedState.Error;
            }
            db.SaveChanges();
            return RedirectToAction("projects", new { /*type = (int)type, */result = (int)(errors.Count > 0 ? NullableFilter.False : NullableFilter.True) });
        }


        /// <summary>
        /// 下载未完成和错误的Project模板
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadProjects(bool? result)
        {
            var list = db.Projects.Where(e => e.City == CurrentUser.City).ToList();

            var workbook = XslHelper.GetWorkbook("templates/自检表.xlsx");

            var sheet = workbook.GetSheetAt(0);
            var rowIndex = 1;

            sheet.InsertRow(rowIndex, list.Count - 1);

            foreach (var item in list)
            {
                var row = sheet.GetRow(rowIndex);
                row.Cells[0].SetCellValue(rowIndex);
                row.Cells[1].SetCellValue(City.浙江省.ToString() + "," + item.City.ToString() + "," + item.County);
                row.Cells[2].SetCellValue(item.ID);
                row.Cells[3].SetCellValue(item.Name);

                rowIndex++;
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            byte[] fileContents = ms.ToArray();
            return File(fileContents, "application/ms-excel", "自检表.xlsx");
        }

        public ActionResult DownloadCoord(string id) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/CoordProject");
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            }
            string filePath = "" ;
            string fileName = "";
            if (id == "ALL")
            {
                string files = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/CoordProject", ((int)CurrentUser.City).ToString());
                
                filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/CoordProject", CurrentUser.City.ToString() + "坐标.zip");
                if (!Fastziping(files, filePath)) {
                    throw new ArgumentException("压缩文件失败");
                }
                fileName = CurrentUser.City.ToString() + "坐标.zip";
            }
            else {
                filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data/CoordProject", ((int)CurrentUser.City).ToString(),id + ".txt");
                fileName = id + ".txt";
            }
            
            if (!System.IO.File.Exists(filePath)) {
                throw new ArgumentException("正确坐标文件不存在");
            }
            
            FileStream fs = new FileStream(filePath, FileMode.Open);
            byte[] bytes = new byte[(int)fs.Length];
            fs.Read(bytes,0,bytes.Length);
            fs.Close();


            Response.Charset = "UTF-8";
            Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8");
            Response.ContentType = "applicaion/octet-stream";

            Response.AddHeader("Content-Disposition", "attachment;filename=" + Server.UrlEncode(fileName));
            Response.BinaryWrite(bytes);
            Response.Flush();
            Response.End();

            return new EmptyResult();;
        }


        public bool Fastziping(string files, string FilePath) {
            try { 
                FastZip zip = new FastZip();
                zip.CreateZip(FilePath,files,true,".txt");
            }catch(Exception er){
                throw new ArgumentException(er.Message);
            }
            return true;
            
        }

    }
}
