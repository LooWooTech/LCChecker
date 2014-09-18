using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{


    public class CheckReport9:CheckEngine,ICheck
    {
       

 
        public CheckReport9(string filePath)
        {
            GetSuppleMessage(filePath);
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
                int Degree = GetDegree(Ship[item].Grade);

                List<int> d = new List<int>();
                for (int j = 1; j < 16; j++)
                {
                    if (j == Degree)
                        continue;
                    d.Add(j);
                }
                int[] D = new int[d.Count()];
                int k = 0;
                foreach (var Blank in d)
                {
                    D[k] = Blank+6;//一等对应第8栏 +7  列是从0开始 -1  =+6；  
                    k++;
                }

                var rule1 = new StringEqual() { ColumnIndex = 3, Data = item };

                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new AndRule()
                    {
                        Rule1 = new StringEqual() { ColumnIndex = 1, Data = Ship[item].City },
                        Rule2 = new StringEqual() { ColumnIndex=2,Data=Ship[item].County}
                    }
                });

                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new StringEqual() { ColumnIndex=4,Data=Ship[item].Name}
                });

                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new DoubleEqual() {ColumnIndex=5,data=Ship[item].AddArea }
                });

                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new MultipleCellRangeRowRule() { ColumnIndices = D, isAny = false, isEmpty = true, isNumeric = false }
                });         
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 22, Values = new[] { "是", "否" } });

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

        /// <summary>
        /// 获取耕地质量等别
        /// </summary>
        /// <param name="grade"></param>
        /// <returns></returns>
        public int GetDegree(string grade)
        {
            double gr;
            double.TryParse(grade, out gr);
            double degree;
            for (int i = 1; i < 16; i++)
            {
                double.TryParse(i.ToString(), out degree);
                if (Math.Abs(degree - gr) < double.Epsilon)
                {
                    return i;
                }
            }
            return 0;
        }



        public bool CheckSpecial(string FilePath,ref string Mistakes,ReportType Type)
        {
            int startRow=0,startCell=0;
            ISheet sheet = OpenSheet(FilePath, true, ref startRow, ref startCell, ref Mistakes,Type);
            if (sheet == null)
            {
                Mistakes = "检索表格内无数据";
                return false;
            }
            startRow++;
            for (int i = startRow; i <= sheet.LastRowNum; i=i+3)
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
                //检查填写等别的数据
                var val = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (val == null)
                    continue;
                if (Ship.ContainsKey(val))
                {
                    Index2 Data = Ship[val];
                    int Degree = GetDegree(Data.Grade);
                    Land rowData = GetExcelLand(sheet, i, Degree + startCell + 6);
                    if (!Data.Land.Compare(rowData)) 
                    {
                        ErrorRow.Add(string.Format("水田：{0}；水浇地：{1}；旱地：{2}", Data.Land.Paddy, Data.Land.Irrigated, Data.Land.Dry));
                    }
                }
                if (ErrorRow.Count() != 0)
                {
                    Error.Add(val, ErrorRow);
                }
                
            }

            
            return true;
        }


        /// <summary>
        /// 表9中获取 水田 水浇地 旱地对应的面积值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="Line">水田行</param>
        /// <param name="Col">所在列</param>
        /// <returns></returns>
        public Land GetExcelLand(ISheet sheet,int Line, int Col)
        {
            Land land = new Land();
            double[] data = new double[3];
            for (int i = 0; i < 3; i++)
            {
                var row = sheet.GetRow(Line++);
                if (row == null)
                    break;
                var value = row.GetCell(Col, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                double Area;
                double.TryParse(value, out Area);
                data[i] = Area;
            }
            land.Paddy = data[0];
            land.Irrigated = data[1];
            land.Dry = data[2];
            return land;
        }


        public  bool Check(string FilePath, ref string Mistakes,ReportType Type)
        {
            return CheckSpecial(FilePath,ref Mistakes,Type);
        }
    }


   
}