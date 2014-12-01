using LCChecker.Areas.Second.Models;
using LCChecker.Helpers;
using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace LCChecker
{
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


        public static ISheet OpenSheet(string filePath, bool findHeader, ref int startRow, ref int startCol, ref string errMsg, SecondReportType Type)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch (Exception er)
            {
                string str = er.ToString();
                errMsg = "打开Excel表格失败：";// +filePath;
                return null;
            }

            if (workbook == null)
            {
                errMsg = "打开Excel表格失败：";// +filePath;
                return null;
            }

            if (workbook.NumberOfSheets == 0)
            {
                errMsg = "Excel文件中没有表格。";
                return null;
            }

            var sheet = workbook.GetSheetAt(0);

            if (findHeader == false) return sheet;

            if (FindHeader(sheet, ref startRow, ref startCol, Type) == false)
            {
                errMsg = "未找到附表文件的表头";
                return null;
            }
            return sheet;
        }
        public static bool FindHeader(this ISheet sheet, ref int startRow, ref int startCell, SecondReportType Type)
        {
            var Name = @"([\w\W])" + Type.GetDescription();
            string[] Header = { "编号", "市", "县" };
            for (int i = 0; i < 20; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < 20; j++)
                    {
                        var value = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (string.IsNullOrEmpty(value))
                            continue;
                        if (value == Type.ToString())
                        {
                            var Row = sheet.GetRow(i + 1);
                            if (Row == null)
                                return false;
                            value = Row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                            if (string.IsNullOrEmpty(value))
                                return false;
                            if (!Regex.IsMatch(value, Name))
                                return false;
                            if (Type == SecondReportType.附表8)
                            {
                                i = i + 5;
                            }
                            else {
                                i = i + 4;
                            }
                            
                            row = sheet.GetRow(i);
                            if (row == null)
                                return false;
                            value = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                            if (string.IsNullOrEmpty(value))
                                return false;
                            if (value == Header[0])
                            {
                                for (int k = 1; k < 3; k++)
                                {
                                    value = row.GetCell(j + k, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                                    if (value != Header[k])
                                    {
                                        return false;
                                    }
                                }
                                startRow = i;
                                startCell = j;
                                return true;
                            }
                            return false;
                        }
                    }
                }
            }
            return false;

        }
        /// <summary>
        /// 获取合并单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowNum"></param>
        /// <param name="colNum"></param>
        /// <param name="rowSpan"></param>
        /// <param name="colSpan"></param>
        /// <returns></returns>
        public static bool isMergeCell(ISheet sheet, int rowNum, int colNum, out int rowSpan, out int colSpan)
        {
            bool result = false;
            rowSpan = 0;
            colSpan = 0;
            if ((rowNum < 1) || (colNum < 1)) return result;
            int rowIndex = rowNum - 1;
            int colIndex = colNum - 1;
            int regionsCount = sheet.NumMergedRegions;
            rowSpan = 1;
            colSpan = 1;
            for (int i = 0; i < regionsCount; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                sheet.IsMergedRegion(range);
                if (range.FirstRow == rowIndex && range.FirstColumn == colIndex)
                {
                    rowSpan = range.LastRow - range.FirstRow + 1;
                    colSpan = range.LastColumn - range.FirstColumn + 1;
                    break;
                }
            }
            try
            {
                result = sheet.GetRow(rowIndex).GetCell(colIndex).IsMergedCell;
            }
            catch
            {
            }
            return result;
        }


        private static Regex _projectIdRe = new Regex(@"^33[0-9]{12}", RegexOptions.Compiled);
        public static bool VerificationID(this string value)
        {
            return _projectIdRe.IsMatch(value);
        }

        public static bool JudgeLand(NPOI.SS.UserModel.ISheet sheet, int Line, int xoffset = 0)
        {
            string[] Lands = new string[] { "水田", "水浇地", "旱地" };
            foreach (var item in Lands)
            {
                IRow row = sheet.GetRow(Line++);
                if (row == null)
                {
                    return false;
                }
                var value = row.GetCell(6 + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    return false;
                if (value != item)
                    return false;
            }
            return true;
        }

        public static bool CheckLand(NPOI.SS.UserModel.ISheet sheet, int Line, ref double LandArea, ref int Degree, ref string Mistakes, int xoffset = 7)
        {
            IRow row = sheet.GetRow(Line);
            if (row == null)
            {
                Mistakes = "未获得相关表格行";
                return false;
            }
            int Max = xoffset + 15;
            bool Flag = false;
            double Area = 0.0;
            for (var i = xoffset; i < Max; i++)
            {
                var cell = row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (string.IsNullOrEmpty(cell.ToString().Trim()))
                    continue;
                if (Flag)
                {
                    Mistakes = "规则2904：水田、水浇地、旱地中同一个分类不允许填写多个质量等别";
                    return false;
                }

                if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                {
                    try
                    {
                        Area = cell.NumericCellValue;
                    }
                    catch
                    {
                        Area = .0;
                    }
                }
                else
                {
                    var val = cell.ToString().Trim();
                    double.TryParse(val, out Area);
                }
                Degree = i - 6;
                Flag = true;
            }
            LandArea = Area;
            //Area = Area / 15;
            //LandArea = Math.Floor(Area * 10000) / 10000;   
            return true;
        }

        public static ICellStyle GetCellStyle(this IWorkbook workbook, XslHeaderStyle str)
        {
            var cellStyle = workbook.CreateCellStyle();

            IFont fontBigheader = workbook.CreateFont();
            fontBigheader.FontHeightInPoints = 16;
            fontBigheader.FontName = "宋体";
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


            IFont font1 = workbook.CreateFont();
            font1.FontName = "宋体";
            font1.FontHeightInPoints = 9;


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
                case XslHeaderStyle.大头:
                    cellStyle.SetFont(fontBigheader);
                    break;
                case XslHeaderStyle.小头:
                    cellStyle.SetFont(fontSmallheader);
                    break;
                case XslHeaderStyle.默认:
                    cellStyle.SetFont(fontText);
                    break;
                case XslHeaderStyle.小小头:
                    cellStyle.SetFont(fontthinheader);
                    break;

                case XslHeaderStyle.文本:
                    cellStyle.SetFont(font1);
                    break;
            }


            return cellStyle;
        }


        public static string GetValue(this ICell cell)
        {
            if (cell == null) return null;
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                default:
                    try
                    {
                        return cell.StringCellValue;
                    }
                    catch
                    {
                        return null;
                    }
            }
        }

        public static Land GetLand(this string value)
        {
            Land land = new Land();
            string val = value.Replace("，", ",").Replace("。", string.Empty);
            var team = val.Split(',');
            foreach (string item in team)
            {
                Regex r = new Regex(@"-?[0-9]");
                string data = r.Match(item).ToString();
                int position = item.IndexOf(data);
                if (position == 0)
                    continue;
                string ground = item.Substring(0, position);
                string area = item.Substring(position);
                double dArea;
                double.TryParse(area, out dArea);
                switch (ground)
                {
                    case "水田":
                        land.Paddy = dArea;
                        break;
                    case "水浇地":
                        land.Irrigated = dArea;
                        break;
                    case "旱地":
                        land.Dry = dArea;
                        break;
                    default: break;
                }
            }
            return land;
        }


        public static IWorkbook CreateExcel(Dictionary<City, Summary> Data,string Name) {
            string[] Header = { "行政区", "项目总数", "通过总数", "失败总数", "未上传数" };
            int[] He = new int[4];
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("sheet1");
            for (var j = 0; j < 5; j++)
            {
                sheet.SetColumnWidth(j, 15 * 256);
            }
            IRow row = sheet.CreateRow(0);
            var cell = row.CreateCell(0);
            cell.CellStyle = workbook.GetCellStyle(XslHeaderStyle.大头);
            cell.SetCellValue(Name);
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

            return workbook;
        }
    }
}