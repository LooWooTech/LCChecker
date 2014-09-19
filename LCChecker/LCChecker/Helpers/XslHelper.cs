using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
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

        public static bool FindHeader(this ISheet sheet, ref int startRow, ref int startCell, ReportType Type)
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
                            i = i + 4;
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
    }
}