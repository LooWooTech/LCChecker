using LCChecker.Rules;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace LCChecker.Models
{
    public class DetectEngine
    {
        private List<RuleInfo> rules = new List<RuleInfo>();
        //private Dictionary<int, Dictionary<int, List<string>>> summaryError = new Dictionary<int, Dictionary<int, List<string>>>();
        //总表中存在的错误
        public Dictionary<string , List<string>> Error = new Dictionary<string , List<string>>();
        //总表中 每行数据对应的行号关系字典
        private Dictionary<string, int> Relatship = new Dictionary<string, int>();

        //private Dictionary<string, int> NameDict = new Dictionary<string, int>();

        public DetectEngine(string region)
        {
            var list = new List<IRowRule>();

            list.Add(new NoLessThanRowRule() { Column1Index = 5, Column2Index = 6 });
            list.Add(new SumRowRule() { SumColumnIndex = 6, ColumnIndices = new[] { 18, 23 } });
            list.Add(new SumRowRule() { SumColumnIndex = 6, ColumnIndices = new[] { 13, 18, 23 } });
            list.Add(new SumRowRule() { SumColumnIndex = 13, ColumnIndices = new[] { 14, 15 } });
            list.Add(new SumRowRule() { SumColumnIndex = 18, ColumnIndices = new[] { 19, 20 } });
            list.Add(new SumRowRule() { SumColumnIndex = 20, ColumnIndices = new[] { 21, 22 } });
            list.Add(new SumRowRule() { SumColumnIndex = 23, ColumnIndices = new[] { 24, 27, 28, 31, 32 } });
            list.Add(new SumRowRule() { SumColumnIndex = 24, ColumnIndices = new[] { 25, 26 } });
            list.Add(new SumRowRule() { SumColumnIndex = 28, ColumnIndices = new[] { 29, 30 } });
            list.Add(new SumRowRule() { SumColumnIndex = 32, ColumnIndices = new[] { 33, 34 } });



            list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" } });
            list.Add(new ConditionalRowRule()
            {
                Condition = new CellEmptyRowRule() { ColumnIndex = 17, isEmpty = false, isNumeric = false },
                Rule = new CellEmptyRowRule() { ColumnIndex = 18, isEmpty = true, isNumeric = true }
            });


            list.Add(new ConditionalRowRule()
            {
                Condition = new CellEmptyRowRule() { ColumnIndex = 17, isEmpty = true, isNumeric = false },
                Rule = new CellEmptyRowRule() { ColumnIndex = 18, isEmpty = false, isNumeric = true }
            });


            list.Add(new UniqueValueRowRule(region) { ColumnIndex = 3, Keyword = "综合整治" });



            //var rule1 = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是" } };

            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 8, isEmpty = true, isNumeric = false }
            //});
            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 9, isEmpty = true, isNumeric = false }
            //});
            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 10, isEmpty = true, isNumeric = false }
            //});

            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 11, isEmpty = true, isNumeric = false }
            //});
            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 12, isEmpty = true, isNumeric = false }
            //});
            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 13, isEmpty = true, isNumeric = true }
            //});
            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 16, isEmpty = true, isNumeric = false }
            //});

            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 19, isEmpty = true, isNumeric = true }
            //});

            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 27, isEmpty = true, isNumeric = false }
            //});

            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 28, isEmpty = true, isNumeric = true }
            //});

            //list.Add(new ConditionalRowRule()
            //{
            //    Condition = rule1,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 31, isEmpty = true, isNumeric = false }
            //});

            //var rule2 = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "否" } };

            //list.Add(new ConditionalRowRule()//第8栏填写：否  那么第9栏 一定要填写（3种类型）  1、调剂出项目对方指标使用有误；2、
            //{
            //    Condition = rule2,
            //    Rule =
            //        new CellRangeRowRule()
            //        {
            //            ColumnIndex = 8,
            //            Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误", "3、属于复垦、整理、综合整治等项目无法确认信息项目" }
            //        }
            //});

            //list.Add(new ConditionalRowRule()//第8栏 填写：  否  第32栏  无填写
            //{
            //    Condition = rule2,
            //    Rule = new CellEmptyRowRule() { ColumnIndex = 31, isEmpty = true, isNumeric = false }
            //});


            //list.Add(new ConditionalRowRule()//第8栏填写：否  第9栏 填写了 1 2类型  那么  11、12、13、17、20、28、32栏中至少有一栏填写
            //{
            //    Condition = rule2,
            //    Rule = new ConditionalRowRule()
            //    {
            //        Condition = new CellRangeRowRule()
            //        {
            //            ColumnIndex = 8,
            //            Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误" }
            //        },
            //        Rule = new MultipleCellRangeRowRule()
            //        {
            //            ColumnIndices = new[] { 10, 11, 12, 16, 19, 27, 31 },
            //            isAny = true,
            //            isEmpty = false,
            //            isNumeric = false
            //        }
            //    }
            //});

            //list.Add(new ConditionalRowRule()//第8栏填写：否   第9栏填写了类型1 那么20栏要有面积
            //{
            //    Condition = rule2,
            //    Rule = new ConditionalRowRule()
            //    {
            //        Condition = new CellRangeRowRule()
            //        {
            //            ColumnIndex = 8,
            //            Values = new[] { "1、调剂出项目对方指标使用有误" }
            //        },
            //        Rule = new CellEmptyRowRule()
            //        {
            //            ColumnIndex = 19,
            //            isEmpty = false,
            //            isNumeric = true
            //        }
            //    }
            //});

            //list.Add(new ConditionalRowRule()//第8栏填写：否 第9栏为类型2  第11栏、12、13、17、28、32 有面积
            //{
            //    Condition = rule2,
            //    Rule = new ConditionalRowRule()
            //    {
            //        Condition = new CellRangeRowRule()
            //        {
            //            ColumnIndex = 8,
            //            Values = new[] { "2、本县自行补充(含尚未调剂出)项目有误" }
            //        },
            //        Rule = new MultipleCellRangeRowRule()
            //        {
            //            ColumnIndices = new[] { 10, 11, 12, 16, 27, 31 },
            //            isAny = false,
            //            isEmpty = false,
            //            isNumeric = true
            //        }
            //    }
            //});

            //list.Add(new ConditionalRowRule()//第8栏填写：否 第9栏为类型3  ；28、29栏有面积 并且24、33栏无面积
            //{
            //    Condition = rule2,
            //    Rule = new ConditionalRowRule()
            //    {
            //        Condition = new CellRangeRowRule()
            //        {
            //            ColumnIndex = 8,
            //            Values = new[] { "3、属于复垦、整理、综合整治等项目无法确认信息项目" }
            //        },
            //        Rule = new AndRule()
            //        {
            //            Rule1 = new MultipleCellRangeRowRule()
            //            {
            //                ColumnIndices = new[] { 27, 28 },
            //                isAny = false,
            //                isEmpty = false,
            //                isNumeric = true
            //            },
            //            Rule2 = new MultipleCellRangeRowRule()
            //            {
            //                ColumnIndices = new[] { 23, 32 },
            //                isAny = false,
            //                isEmpty = true,
            //                isNumeric = true
            //            }
            //        }
            //    }
            //});

            //list.Add(new CellEmptyRowRule() { ColumnIndex=36,isEmpty=false,isNumeric=false});

            list.Add(new Format() { ColumnIndex = 35, form = "0.0" });


            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }

        }
        /*
         * 检查该区域的总表格 错误信息保存在Error中
         */
        public bool CheckSummaryExcel(string summaryPath,out string mistakes)
        {
            XSSFWorkbook summWorkbook;
            try
            {
                FileStream fs = new FileStream(summaryPath, FileMode.Open, FileAccess.Read);
                summWorkbook = new XSSFWorkbook(fs);
            }
            catch (Exception er){
                mistakes = er.Message;
                return false;
            }
            ISheet sheet = summWorkbook.GetSheetAt(0);
            
            int MaxRow = sheet.LastRowNum;
            for (int i = 0,h=1; i <= MaxRow; i++,h++)
            {
                List<string> ErrorRow = new List<string>();
                IRow Row = sheet.GetRow(h);
                if (Row == null)
                    continue;
                var value = Row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                Relatship.Add(value, h);

                foreach (var item in rules)
                {
                    if (!item.Rule.Check(Row))
                    {
                        ErrorRow.Add(item.Rule.Name);
                    }
                }
                if (ErrorRow.Count() != 0)
                {
                    Error.Add(value, ErrorRow);
                }

            }
            mistakes = "";
            return true;
        }

        /*
         检查提交的表格
         * submitPath 提交表格路径  summaryPath 总表路径
         * 用途：将总表错误 提交中正确的行 数据更新到总表
         * subError 字典 key 是提交表格中错误行数据的行号（不需要加1 在NPOI中）
         * 成功检索数据  返回true ；否则为false
         */
        public bool CheckSubmitExcel(string SubmitPath,string summaryPath,string resultPath,ref Dictionary<string,List<string>> subError,out string mistakes)
        {
            XSSFWorkbook summaryBook;
            IWorkbook workbook;
            try
            {
                FileStream summ = new FileStream(summaryPath, FileMode.Open, FileAccess.ReadWrite);
                summaryBook = new XSSFWorkbook(summ);
            }
            catch (Exception er){
                mistakes = er.Message;
                return false;
            }
            try
            {
                FileStream fs = new FileStream(SubmitPath, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
            }
            catch (Exception er){
                mistakes = er.Message;
                return false;
            }
            ISheet summSheet = summaryBook.GetSheetAt(0);

            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            if (!FindHeader(sheet, ref startRow, ref startCell))
            {
                subError.Add("", new List<string>() { "未找到表头,导致未能检索表格数据" });
                mistakes = "未找到表头，导致未能检索表格数据";
                return false;
            }
            startRow++;
            IRow Row;
            int MaxRow = sheet.LastRowNum;
            for (int irow = startRow; irow <= MaxRow; irow++)
            {
                Row = sheet.GetRow(irow);
                bool flag = false;
                List<string> ErrorRow = new List<string>();
                foreach (var item in rules)
                {
                    if (!item.Rule.Check(Row, startCell))
                    {
                        ErrorRow.Add(item.Rule.Name);
                        flag = true;
                    }
                }
                var value = Row.GetCell(startCell + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (flag)
                {
                    subError.Add(value, ErrorRow);
                }
                else{
                    if(Error.ContainsKey(value))
                    {
                        int ErrorNumer=Relatship[value];
                        IRow errorRow=summSheet.GetRow(ErrorNumer);
                        int MaxCell=startCell+43;
                        for (int i = startCell,j=0; i <= MaxCell; i++,j++)
                        {
                            var CorrectCell = Row.GetCell(i);
                            if (CorrectCell == null)
                                continue;
                            var ErrorCell = errorRow.GetCell(j);
                            if (ErrorCell == null)
                                ErrorCell = errorRow.CreateCell(j, CorrectCell.CellType);
                            switch (CorrectCell.CellType)
                            {
                                case CellType.Boolean:
                                    ErrorCell.SetCellValue(CorrectCell.BooleanCellValue);
                                    break;
                                case CellType.Numeric:
                                    ErrorCell.SetCellValue(CorrectCell.NumericCellValue);
                                    break;
                                case CellType.String:
                                    ErrorCell.SetCellValue(CorrectCell.StringCellValue);
                                    break;
                            }
                        }
                            Error.Remove(value);
                    }
                }
            }
            try
            {
                FileStream fs = new FileStream(resultPath, FileMode.Create, FileAccess.Write);
                summaryBook.Write(fs);
                FileStream summfs = new FileStream(summaryPath, FileMode.Open, FileAccess.Write);
                summaryBook.Write(summfs);
            }
            catch (Exception er){
                mistakes = er.Message;
                return false;
            }
            mistakes = "";
            return true;
        }



        /*查找表格的开始 各行的开始*/
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
    }
}