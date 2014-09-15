using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckReport2:CheckEngine
    {
        public Dictionary<string, Index2> Ship = new Dictionary<string, Index2>();

        public CheckReport2(string filePath)
        {
            GetMessage(filePath);
            var list = new List<IRowRule>();
            list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" } });
            foreach (var item in Ship.Keys)
            {
                var rule1 = new StringEqual() { ColumnIndex = 3, Data = item };
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new DoubleEqual() { ColumnIndex = 5, data = Ship[item].AddArea }
                });

                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new DoubleEqual() { ColumnIndex = 6, data = Ship[item].Indicators }
                });

            }

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

        public bool GetMessage(string filePath)
        {
            IWorkbook workbook = null;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            if (workbook == null)
            {
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int StartRow = 1;
            IRow row = sheet.GetRow(StartRow++);
            while (row != null)
            {
                var value = row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (value == "")
                {
                    break;
                }
                var value1 = row.GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                var value2 = row.GetCell(6, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                double a = double.Parse(value1);
                double b = double.Parse(value2);
                Index2 one = new Index2() { AddArea = a, Indicators = b };
                Ship.Add(value, one);
                row = sheet.GetRow(StartRow++);
            }
            return true;
        }

    }
}