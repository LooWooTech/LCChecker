using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckReport9
    {
        private string[] LandType={"水田","水浇田","旱地"};

        public List<RuleInfo> rules = new List<RuleInfo>();
        public Dictionary<string, Index2> Ship = new Dictionary<string, Index2>();
        public Dictionary<string,double>  Grade=new Dictionary<string,double>();

        //public Dictionary<string,double> 
        public CheckReport9()
        {
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
                    Rule = new StringEqual() { ColumnIndex = 1, Data = Ship[item].City }
                });

                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new StringEqual() { ColumnIndex = 2, Data = Ship[item].County }
                });
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new StringEqual() { ColumnIndex = 4, Data = Ship[item].Name }
                });
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new DoubleEqual() { ColumnIndex = 5, data = Ship[item].AddArea }
                });
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 10, Values = new[] { "是", "否" } });
            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }
        public void GetMessage(string filePath)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch {
                return;
            }
            
            if (workbook == null)
            {
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

            for (int i = 1; i < +sheet.LastRowNum; i++)
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
                double data1 = double.Parse(AddArea);
                if (cityName.Length < 3)
                {
                    continue;
                }
                Ship.Add(value, new Index2
                {
                    City = cityName[1],
                    County = cityName[2],
                    Name = value1,
                    AddArea = data1
                });
                var val = row.GetCell(36, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                var val2 = val.Replace("，", ",").Replace(".", string.Empty).Replace("。", string.Empty);

            }



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

        public bool CheckSpecial(string FilePath,ref string Mistakes)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch {
                Mistakes = "打开文件失败";
                return false;
            }
            if (workbook == null)
            {
                Mistakes = "获取文件信息失败";
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            if (!FindHeader(sheet, ref startRow, ref startCell))
            {
                Mistakes = "未找到文件的表头[编号 市 县]";
                return false;
            }
            startRow++;
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                List<string> ErrorRow = new List<string>();
                foreach (var item in rules)
                {
                    if (!item.Rule.Check(row, startCell))
                    {
                        ErrorRow.Add(item.Rule.Name);
                    }
                }
                
            }

            
            return true;
        }
    }
}