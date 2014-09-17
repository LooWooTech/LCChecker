using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckReport7 : CheckEngine
    {


        public CheckReport7(string filePath)
        {
            GetMessage(filePath);
            var list = new List<IRowRule>();
            int count = Ship.Count();
            string[] IDS = new string[count];
            int i = 0;
            foreach (var item in Ship.Keys)
            {
                IDS[i] = item;
                i++;
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 3, Values = IDS });
            foreach (var item in Ship.Keys)
            {
                var rule1 = new StringEqual() { ColumnIndex = 3, Data = item };
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new AndRule()
                    {
                        Rule1 = new AndRule()
                        {
                            Rule1 = new AndRule()
                            {
                                Rule1 = new StringEqual() { ColumnIndex = 1, Data = Ship[item].City },
                                Rule2 = new StringEqual() { ColumnIndex = 2, Data = Ship[item].County }
                            },
                            Rule2 = new AndRule()
                            {
                                Rule1 = new StringEqual() { ColumnIndex = 4, Data = Ship[item].Name },
                                Rule2 = new DoubleEqual() { ColumnIndex=5,data=Ship[item].AddArea}
                            }
                        },
                        Rule2 = new StringEqual() { ColumnIndex = 7, Data = Ship[item].Grade }
                    }
                });
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" } });

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }


        public new void GetMessage(string filePath)
        {
            string fault = "";
            IWorkbook workbook = GetWorkbook(filePath, ref fault);
            if (workbook == null)
            {
                Error.Add("大错误", new List<string>() { "获取项目总表失败" });
                return;
            }
            RuleInfo engine = new RuleInfo()//要验证项目是存在补充耕地面积 那么第14 19 24 栏至少有一栏是有面积的 假如这个条件成立那么就是补充耕地项目编号
            {
                Rule = new MultipleCellRangeRowRule()
                {
                    ColumnIndices = new[] { 13, 18, 23 },
                    isAny = true,
                    isEmpty = false,
                    isNumeric = true
                }
            };
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 1;
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                if (!engine.Rule.Check(row))
                    continue;
                var value = row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    break;
                var value1 = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                var cityName = row.Cells[1].StringCellValue.Split(',');
                var AddArea = row.Cells[5].StringCellValue;
                var grade = row.Cells[35].StringCellValue;
                double data = double.Parse(AddArea);
                if (cityName.Length < 3)
                {
                    continue;
                }
                Ship.Add(value, new Index2
                {
                    City = cityName[1],
                    County = cityName[2],
                    Name = value1,
                    AddArea = data,
                    Grade = grade
                });
            }
        }
    }
}