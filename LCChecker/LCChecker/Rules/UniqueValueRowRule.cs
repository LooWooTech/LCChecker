using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Rules
{
    internal class UniqueValueRowRule:IRowRule
    {
        public int ColumnIndex { get; set; }
        public IRowRule Rule { get; set; }
        public string Keyword { get; set; }
       // public string summary { get; set; }
        private Dictionary<string, int> nameDict = new Dictionary<string, int>();
        public UniqueValueRowRule(string summary)
        {
            try
            {
                FileStream fs = new FileStream(summary, FileMode.Open, FileAccess.Read);
                XSSFWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                int all = sheet.LastRowNum;
                for (int i = 0; i <= all; i++)
                {
                    IRow row = sheet.GetRow(i);
                    var value = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                    if (nameDict.ContainsKey(value))
                    {
                        nameDict[value] = nameDict[value] + 1;
                    }
                    else {
                        nameDict.Add(value, 1);
                    }
                }
            }
            catch { 
                
            }
        }


        public string Name {
            get { return "项目名称包含“综合整治”需复核重复备案需删除"; }
        }

        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value1 = row.GetCell(xoffset + ColumnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();

            if (!string.IsNullOrEmpty(Keyword) && value1.Contains(Keyword))
            {
                if (!nameDict.ContainsKey(value1))
                {
                    return false;
                }
                if (nameDict[value1] != 1)
                {
                    return false;
                }
                else {
                    return true;
                }
            }
            return true;
        }

        
    }
}