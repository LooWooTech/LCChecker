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


        public bool Check(string filePath,ref string mistakes)
        {
            IWorkbook workbook = null;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            if (workbook == null)
            {
                mistakes = "打开文件失败";
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            if (!FindHeader(sheet, ref startRow, ref startCell))
            {
                mistakes = "未找到文件表头：" + filePath;
                return false;
            }
            startRow++;
            var row = sheet.GetRow(startRow++);
            while (row != null)
            {
                List<string> ErrorRow = new List<string>();
                var value = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (value == "")
                    break;
                foreach (var item in rules)
                {
                    if (!item.Rule.Check(row, 0))
                    {
                        ErrorRow.Add(item.Rule.Name);
                    }
                }
                if (ErrorRow.Count() != 0)
                {
                    Error.Add(value, ErrorRow);
                }
                row = sheet.GetRow(startRow++);
            }
            return true;
        }

        public bool FindHeader(ISheet sheet, ref int startRow, ref int startCell)
        {
            string[] Header = { "编号", "市", "县" };
            for (int i = 0; i < 10; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < 10; j++)
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
    }
}