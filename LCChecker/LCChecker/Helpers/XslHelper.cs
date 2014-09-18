using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker
{
    //public enum XslHeaderStyle
    //{
    //    大头,
    //    小头,
    //    小小头,
    //    文本,
    //    默认
    //}
    
    public static class XslHelper
    {
        public static IWorkbook GetWorkbook(HttpPostedFileBase file)
        {
            var ext = Path.GetExtension(file.FileName);
            if (ext == ".xls")
            {
                return new HSSFWorkbook(file.InputStream);
            }
            else
            {
                return new XSSFWorkbook(file.InputStream);
            }
        }

        public static IWorkbook GetWorkbook(string path)
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return WorkbookFactory.Create(fs);
            }
        }

        public static void InsertRow(this ISheet sheet, int startRowIndex, int count)
        {
            sheet.ShiftRows(startRowIndex + 1, sheet.LastRowNum, count, true, false);
            var templateRow = sheet.GetRow(startRowIndex);
            for (var rowIndex = startRowIndex + 1; rowIndex < startRowIndex + count + 1; rowIndex++)
            {
                var row = sheet.CreateRow(rowIndex);
                if (templateRow.RowStyle != null)
                {
                    row.RowStyle = templateRow.RowStyle;
                }
                row.Height = templateRow.Height;
                for (var cellIndex = 0; cellIndex < templateRow.Cells.Count; cellIndex++)
                {
                    var cellTemplate = templateRow.Cells[cellIndex];
                    var cell = row.CreateCell(cellIndex, cellTemplate.CellType);
                    if (cellTemplate.CellStyle != null)
                    {
                        cell.CellStyle = cellTemplate.CellStyle;
                    }
                }
            }
        }

        //public static ICellStyle GetCellStyle(this IWorkbook workbook, XslHeaderStyle str)
        //{
        //    var cellStyle = workbook.CreateCellStyle();

        //    IFont fontBigheader = workbook.CreateFont();
        //    fontBigheader.FontHeightInPoints = 22;
        //    fontBigheader.FontName = "微软雅黑";
        //    fontBigheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

        //    IFont fontSmallheader = workbook.CreateFont();
        //    fontSmallheader.FontHeightInPoints = 14;
        //    fontSmallheader.FontName = "黑体";

        //    IFont fontText = workbook.CreateFont();
        //    fontText.FontHeightInPoints = 12;
        //    fontText.FontName = "宋体";

        //    IFont fontthinheader = workbook.CreateFont();
        //    fontthinheader.FontName = "宋体";
        //    fontthinheader.FontHeightInPoints = 11;
        //    fontthinheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;


        //    IFont font1 = workbook.CreateFont();
        //    font1.FontName = "宋体";
        //    font1.FontHeightInPoints = 9;


        //    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

        //    //边框颜色
        //    cellStyle.BottomBorderColor = HSSFColor.OliveGreen.Black.Index;
        //    cellStyle.TopBorderColor = HSSFColor.OliveGreen.Black.Index;

        //    //背景图形
        //    cellStyle.FillForegroundColor = HSSFColor.White.Index;
        //    cellStyle.FillBackgroundColor = HSSFColor.Black.Index;

        //    //文本对齐  左对齐  居中  右对齐  现在是居中
        //    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

        //    //垂直对齐
        //    cellStyle.VerticalAlignment = VerticalAlignment.Center;

        //    //自动换行
        //    cellStyle.WrapText = true;

        //    //缩进
        //    cellStyle.Indention = 0;

        //    switch (str)
        //    {
        //        case XslHeaderStyle.大头:
        //            cellStyle.SetFont(fontBigheader);
        //            break;
        //        case XslHeaderStyle.小头:
        //            cellStyle.SetFont(fontSmallheader);
        //            break;
        //        case XslHeaderStyle.默认:
        //            cellStyle.SetFont(fontText);
        //            break;
        //        case XslHeaderStyle.小小头:
        //            cellStyle.SetFont(fontthinheader);
        //            break;

        //        case XslHeaderStyle.文本:
        //            cellStyle.SetFont(font1);
        //            break;
        //    }


        //    return cellStyle;
        //}
    }
}