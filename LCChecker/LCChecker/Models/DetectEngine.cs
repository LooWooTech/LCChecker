using LCChecker.Rules;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

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
        //public int submit;

        public DetectEngine()
        {
            var list = new List<IRowRule>();

            list.Add(new NoLessThanRowRule() { Column1Index = 5, Column2Index = 6 });
            list.Add(new SumRowRule() { SumColumnIndex = 6, ColumnIndices = new[] { 18, 23 } });
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

            list.Add(new UniqueValueRowRule() { ColumnIndex = 3, Keyword = "综合整治" });

            var rule1 = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是" } };

            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 8, isEmpty = true, isNumeric = false }
            });
            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 9, isEmpty = true, isNumeric = false }
            });
            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 10, isEmpty = true, isNumeric = false }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 11, isEmpty = true, isNumeric = false }
            });
            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 12, isEmpty = true, isNumeric = false }
            });
            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 13, isEmpty = true, isNumeric = true }
            });
            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 16, isEmpty = true, isNumeric = false }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 19, isEmpty = true, isNumeric = true }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 27, isEmpty = true, isNumeric = false }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 28, isEmpty = true, isNumeric = true }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule1,
                Rule = new CellEmptyRowRule() { ColumnIndex = 31, isEmpty = true, isNumeric = false }
            });

            var rule2 = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "否" } };

            list.Add(new ConditionalRowRule()
            {
                Condition = rule2,
                Rule =
                    new CellRangeRowRule()
                    {
                        ColumnIndex = 8,
                        Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误", "3、属于复垦、整理、综合整治等项目无法确认信息项目" }
                    }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule2,
                Rule = new CellEmptyRowRule() { ColumnIndex = 31, isEmpty = true, isNumeric = false }
            });


            list.Add(new ConditionalRowRule()
            {
                Condition = rule2,
                Rule = new ConditionalRowRule()
                {
                    Condition = new CellRangeRowRule()
                    {
                        ColumnIndex = 8,
                        Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误" }
                    },
                    Rule = new MultipleCellRangeRowRule()
                    {
                        ColumnIndices = new[] { 10, 11, 12, 16, 19, 27, 31 },
                        isAny = true,
                        isEmpty = false,
                        isNumeric = false
                    }
                }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule2,
                Rule = new ConditionalRowRule()
                {
                    Condition = new CellRangeRowRule()
                    {
                        ColumnIndex = 8,
                        Values = new[] { "1、调剂出项目对方指标使用有误" }
                    },
                    Rule = new CellEmptyRowRule()
                    {
                        ColumnIndex = 19,
                        isEmpty = false,
                        isNumeric = true
                    }
                }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule2,
                Rule = new ConditionalRowRule()
                {
                    Condition = new CellRangeRowRule()
                    {
                        ColumnIndex = 8,
                        Values = new[] { "2、本县自行补充(含尚未调剂出)项目有误" }
                    },
                    Rule = new MultipleCellRangeRowRule()
                    {
                        ColumnIndices = new[] { 10, 11, 12, 16, 27, 31 },
                        isAny = false,
                        isEmpty = false,
                        isNumeric = true
                    }
                }
            });

            list.Add(new ConditionalRowRule()
            {
                Condition = rule2,
                Rule = new ConditionalRowRule()
                {
                    Condition = new CellRangeRowRule()
                    {
                        ColumnIndex = 8,
                        Values = new[] { "3、属于复垦、整理、综合整治等项目无法确认信息项目" }
                    },
                    Rule = new AndRule()
                    {
                        Rule1 = new MultipleCellRangeRowRule()
                        {
                            ColumnIndices = new[] { 27, 28 },
                            isAny = false,
                            isEmpty = false,
                            isNumeric = true
                        },
                        Rule2 = new MultipleCellRangeRowRule()
                        {
                            ColumnIndices = new[] { 23, 32 },
                            isAny = false,
                            isEmpty = true,
                            isNumeric = true
                        }
                    }
                }
            });


            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }

        }
        /*
         * 检查该区域的总表格 错误信息保存在Error中
         */
        public bool CheckSummaryExcel(string summaryPath)
        {
            try
            {
                FileStream fs = new FileStream(summaryPath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook summWorkbook = new XSSFWorkbook(fs);
                ISheet sheet = summWorkbook.GetSheetAt(0);
                int h = 1;
                IRow Row = sheet.GetRow(h);
                while (Row != null)
                {
                    List<string> ErrorRow = new List<string>();
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
                    h++;
                    Row = sheet.GetRow(h);
                }

            }
            catch {
                return false;
            }
            return true;
        }

        /*
         检查提交的表格
         * submitPath 提交表格路径  summaryPath 总表路径
         * 用途：将总表错误 提交中正确的行 数据更新到总表
         * subError 字典 key 是提交表格中错误行数据的行号（不需要加1 在NPOI中）
         * 成功检索数据  返回true ；否则为false
         */
        public bool CheckSubmitExcel(string SubmitPath,string summaryPath,ref Dictionary<string,List<string>> subError)
        {
            try
            {
                FileStream summ = new FileStream(summaryPath, FileMode.Open, FileAccess.ReadWrite);
                XSSFWorkbook summaryBook = new XSSFWorkbook(summ);
                ISheet summSheet = summaryBook.GetSheetAt(0);
                try
                {
                    FileStream fs = new FileStream(SubmitPath, FileMode.Open, FileAccess.Read);
                    XSSFWorkbook workbook = new XSSFWorkbook(fs);
                    ISheet sheet = workbook.GetSheetAt(0);
                    int startRow = 0, startCell = 0;
                    if (!FindHeader(sheet, ref startRow, ref startCell))
                    {
                        subError.Add("", new List<string>() { "未找到表头,导致未能检索表格数据" });
                        return false;
                    }
                    startRow++;
                    IRow Row ;
                    int Rowsum = sheet.LastRowNum;

                    for (int irow = startRow; irow <= Rowsum; irow++)
                    {
                        Row = sheet.GetRow(irow);
                        bool flag = false;
                        List<string> ErrorRow = new List<string>();

                        foreach (var item in rules)
                        {
                            if (!item.Rule.Check(Row))
                            {
                                ErrorRow.Add(item.Rule.Name);
                                flag = true;
                            }
                        }
                        var value = Row.GetCell(startCell + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        //有错误 对于提交的表格的某一行有错误，那么将错误信息保存在
                        if (flag)
                        {
                            //第几行 错误列表
                            subError.Add(value, ErrorRow);
                        }
                        //没有错误
                        else
                        {
                            //查看总表是否有错误记录
                            if (Error.ContainsKey(value))
                            {
                                int rowNumber = Relatship[value];
                                IRow changeRow = summSheet.GetRow(rowNumber);
                                int j = 0;
                                for (int i = startCell; i <= Row.LastCellNum; i++)
                                {
                                    //获取提交表格的 Row行的数据
                                    var Correctvalue = Row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                                    //changRow是总表相应更新的行
                                    changeRow.CreateCell(j).SetCellValue(Correctvalue);
                                }
                                //将正确的数据更新到总表之后，那么删除该行在错误字典中的信息
                                Error.Remove(value);
                            }
                        }
                    }
                }
                catch {
                    return false;
                }
            }
            catch {
                return false;
            }
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