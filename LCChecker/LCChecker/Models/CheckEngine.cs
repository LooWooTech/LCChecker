using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckEngine
    {
        public List<RuleInfo> rules = new List<RuleInfo>();

        public Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();

        public Dictionary<string, Index2> Ship = new Dictionary<string, Index2>();

        public bool Check(string filePath, ref string mistakes)
        {
            IWorkbook workbook = GetWorkbook(filePath, ref mistakes);
            if (workbook == null)
                return false;
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            if (!FindHeader(sheet, ref startRow, ref startCell))
            {
                mistakes = "未找到文件表头[编号 市 县]：" + filePath;
                return false;
            }
            startRow++;
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    break;
                List<string> ErrorRow = new List<string>();
                var value = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                foreach (var item in rules)
                {
                    if (!item.Rule.Check(row, startCell))
                    {
                        ErrorRow.Add(item.Rule.Name);
                    }
                }
                if (ErrorRow.Count() != 0)
                {
                    Error.Add(value, ErrorRow);
                }
            }

            return true;
        }

        public bool FindHeader(ISheet sheet, ref int startRow, ref int startCell)
        {
            string[] Header = { "编号", "市", "县" };
            for (int i = 0; i < 20; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < 20; j++)
                    {
                        var value = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
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
                    }
                }
            }
            return false;

        }

        public IWorkbook GetWorkbook(string filePath, ref string Mistakes)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch
            {
                Mistakes = "打开文件失败";
                return null;
            }
            if (workbook == null)
            {
                Mistakes = "打开文件失败";
                return null;
            }
            return workbook;
        }

        public void GetMessage(string filePath)
        {
            string fault = "";
            IWorkbook workbook = GetWorkbook(filePath, ref fault);
            if (workbook == null)
            {
                Error.Add("大错误", new List<string>() { "获取项目总表失败" });
                return;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 1;
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                var value = row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    break;
                var value1 = row.Cells[3].StringCellValue;
                var cityName = row.Cells[1].StringCellValue.Split(',');
                if (cityName.Length < 3)
                {
                    continue;
                }
                Ship.Add(value, new Index2
                {
                    City = cityName[1],
                    County = cityName[2],
                    Name = value1
                });
            }
        }
    }
}