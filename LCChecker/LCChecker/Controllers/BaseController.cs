using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    public class BaseController : Controller
    {
        private LCDbContext db = new LCDbContext();

        /*
         * 数据采集  
         * 采集数据种类：项目个数 flag 用于判断是否第一次采集 /更新数据
         */
        public bool CollectData(string region,ref string errorInformation)
        {
            Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();
            Dictionary<string, int> ship = new Dictionary<string, int>();
            string DataPath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region), "summary.xlsx");
            //检查该表中存在错误的行
            DetectEngine Engine = new DetectEngine(DataPath);
            string ErrorMessage="";
            if (!Engine.CheckExcel(DataPath, ref ErrorMessage, ref Error, ref ship))
            {
                return false;
            }
            Detect record = db.DETECT.Where(x => x.region == region).FirstOrDefault();
            if (record == null)
            {
                errorInformation = "未找到相关记录";
                return false;
            }
            //ship中不包括表头 所以ship的个数就是项目个数
            //error中就是错误的个数，修改正确的个数 那就是ship.count（）-error.count（）
            if (record.sum != ship.Count())
            {
                record.sum = ship.Count();
            }

            record.Correct = ship.Count() - Error.Count();
            if (ModelState.IsValid)
            {
                db.Entry(record).State = EntityState.Modified;
                db.SaveChanges();
            }
            return true;
        }

        public enum stylexls
        {
            大头,
            小头,
            小小头,
            默认
        }
        /*
         * 设置单元格  格式
         */
        public static ICellStyle GetCellStyle(IWorkbook workbook, stylexls str)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            IFont fontBigheader = workbook.CreateFont();
            fontBigheader.FontHeightInPoints = 22;
            fontBigheader.FontName = "微软雅黑";
            fontBigheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            IFont fontSmallheader = workbook.CreateFont();
            fontSmallheader.FontHeightInPoints = 14;
            fontSmallheader.FontName = "黑体";

            IFont fontText = workbook.CreateFont();
            fontText.FontHeightInPoints = 12;
            fontText.FontName = "宋体";

            IFont fontthinheader = workbook.CreateFont();
            fontthinheader.FontName = "宋体";
            fontthinheader.FontHeightInPoints = 11;
            fontthinheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;


            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            //边框颜色
            cellStyle.BottomBorderColor = HSSFColor.OliveGreen.Black.Index;
            cellStyle.TopBorderColor = HSSFColor.OliveGreen.Black.Index;

            //背景图形
            cellStyle.FillForegroundColor = HSSFColor.White.Index;
            cellStyle.FillBackgroundColor = HSSFColor.Black.Index;

            //文本对齐  左对齐  居中  右对齐  现在是居中
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            //垂直对齐
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            //自动换行
            cellStyle.WrapText = true;

            //缩进
            cellStyle.Indention = 0;

            switch (str)
            {
                case stylexls.大头:
                    cellStyle.SetFont(fontBigheader);
                    break;
                case stylexls.小头:
                    cellStyle.SetFont(fontSmallheader);
                    break;
                case stylexls.默认:
                    cellStyle.SetFont(fontText);
                    break;
                case stylexls.小小头:
                    cellStyle.SetFont(fontthinheader);
                    break;
            }


            return cellStyle;
        }

        /*用途：当用户第一次上传表格的时候，保存总表  文件拷贝好像有自己的函数 ，貌似总是出问题 
         */
        public bool Copy(string source, string reborn,ref string mistakes)
        {
            IWorkbook workbook;
            try {
                FileStream fs = new FileStream(source, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
                fs.Close();
            }
            catch(Exception er)
            {
                mistakes = er.Message;
                return false;
            }

            try
            {
                FileStream fs = new FileStream(reborn, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                mistakes = er.Message;
                return false;
            }
            return true;
        }

        
    }
}
