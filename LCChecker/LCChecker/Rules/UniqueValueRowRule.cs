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
            IWorkbook workbook = null;
            //string SummaryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, summary);
            try
            {
                FileStream fs = new FileStream(summary, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
                fs.Close();
            }
            catch
            {
 
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            FindHeader(sheet, ref startRow, ref startCell);
            startRow++;
            int LastNumber = sheet.LastRowNum;
            for (int y = startRow; y <= LastNumber; y++)
            {
                IRow row = sheet.GetRow(y);
                var value = row.GetCell(startCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (nameDict.ContainsKey(value))
                {
                    nameDict[value]=nameDict[value]+1;
                }
                else{
                    nameDict.Add(value,1);
                }
            } 
        }
        private bool FindHeader(ISheet sheet, ref int startrow, ref int startcol)
        {
            for (int i = 0; i < 5; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        var value = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (value == "1栏")
                        {
                            for (int k = 0; k < 43; k++)
                            {
                                value = row.GetCell(k + j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                                if (value != string.Format("{0}栏", k + 1))
                                {
                                    return false;
                                }
                            }
                            startrow = i;
                            startcol = j;
                            return true;
                        }
                    }
                }
            }
            return false;
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