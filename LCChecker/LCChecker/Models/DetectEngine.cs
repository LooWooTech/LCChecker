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

        public Dictionary<string, List<string>> summaryError = new Dictionary<string, List<string>>();

        public Dictionary<string, List<string>> subError = new Dictionary<string, List<string>>();

        public Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();




        public DetectEngine(string region)
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


            list.Add(new UniqueValueRowRule(region) { ColumnIndex = 3, Keyword = "综合整治" });



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

            list.Add(new ConditionalRowRule()//第8栏填写：否  那么第9栏 一定要填写（3种类型）  1、调剂出项目对方指标使用有误；2、
            {
                Condition = rule2,
                Rule =
                    new CellRangeRowRule()
                    {
                        ColumnIndex = 8,
                        Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误", "3、属于复垦、整理、综合整治等项目无法确认信息项目" }
                    }
            });

            list.Add(new ConditionalRowRule()//第8栏 填写：  否  第32栏  无填写
            {
                Condition = rule2,
                Rule = new CellEmptyRowRule() { ColumnIndex = 31, isEmpty = true, isNumeric = false }
            });


            list.Add(new ConditionalRowRule()//第8栏填写：否  第9栏 填写了 1 2类型  那么  11、12、13、17、20、28、32栏中至少有一栏填写
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

            list.Add(new ConditionalRowRule()//第8栏填写：否   第9栏填写了类型1 那么20栏要有面积
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

            list.Add(new ConditionalRowRule()//第8栏填写：否 第9栏为类型2  第11栏、12、13、17、28、32 有面积
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

            list.Add(new ConditionalRowRule()//第8栏填写：否 第9栏为类型3  ；28、29栏有面积 并且24、33栏无面积
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

            //项目类型为2007年前土地整理、土地复垦等项目  第19栏、第28栏、第29栏至少有一栏填写
            list.Add(new ConditionalRowRule()
            {
                Condition = new ProjectType() { Time = 2007, Variety = new[] { "整理", "复垦" } },
                Rule = new MultipleCellRangeRowRule()
                {
                    ColumnIndices = new[] { 18, 27, 28 },
                    isAny = true,
                    isEmpty = false,
                    isNumeric = false
                }
            });
            //项目为2007以后验收土地整理项目  第9栏只能填写 1 ，2 类型
            list.Add(new ConditionalRowRule()
            {
                Condition = new Arrange() { Time = 2007, Variety = "整理" },
                Rule = new CellRangeRowRule()
                {
                    ColumnIndex = 8,
                    Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误" }
                }
            });


            list.Add(new Format() { ColumnIndex = 35, form = "0.0" });

            list.Add(new Special() { ColumnIndex = 36 });




            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }

        }

        public static void InsertRow(ISheet sheet, int startRowIndex, int count)
        {
            sheet.ShiftRows(startRowIndex, sheet.LastRowNum, count, true, false);
            var templateRow = sheet.GetRow(startRowIndex + count);
            for (var rowIndex = startRowIndex; rowIndex < startRowIndex + count; rowIndex++)
            {
                var row = sheet.CreateRow(rowIndex);
                if (templateRow.RowStyle != null)
                {
                    row.RowStyle = templateRow.RowStyle;
                }
                row.Height = templateRow.Height;

                for (var cellIndex = 0; cellIndex < templateRow.Cells.Count; cellIndex++)
                {
                    var cellTemplate = templateRow.Cells[cellIndex];
                    var cell = row.CreateCell(cellIndex, cellTemplate.CellType);
                    cell.SetCellValue(cellTemplate.StringCellValue);
                    if (cellTemplate.CellStyle != null)
                    {
                        cell.CellStyle = cellTemplate.CellStyle;
                    }
                }
            }
        }


        /// <summary>
        /// 确保每个项目是唯一的  获得项目编号 以及项目的个数
        /// </summary>
        /// <param name="filePath">Excel文件</param>
        /// <param name="mistakes">错误信息</param>
        /// <param name="ship">项目编号  个数</param>
        /// <returns></returns>
        public bool GetProject(string filePath, ref string mistakes, ref Dictionary<string, int> ship, ref int LastRowNumber)
        {
            IWorkbook workbook = null;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            if (workbook == null)
            {
                mistakes = "打开文件失败：" + filePath;
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(1);
            int i = 1;
            while (row != null)
            {
                var value = row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (value != "")
                {
                    if (ship.ContainsKey(value))
                    {
                        ship[value]++;
                    }
                    else
                    {
                        ship.Add(value, 1);
                    }
                }
                row = sheet.GetRow(++i);
            }
            LastRowNumber = i;
            return true;
        }


        /*检查表格中的错误
         * Path 检查表格的路径
         * mistakes 检查过程可能失败、出错信息
         * ErrorMessage 检查表格中的错误
         *  relatship 表格中每行中的编号对应的行号关系
         */
        public bool CheckExcel(string Path, ref string mistakes, ref Dictionary<string, List<string>> ErrorMessage, ref Dictionary<string, int> relatship)
        {
            IWorkbook workbook;
            try
            {
                FileStream fs = new FileStream(Path, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                mistakes = er.Message;
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            if (!FindHeader(sheet, ref startRow, ref startCell))
            {
                mistakes = "未找到表格：" + Path + "的表头";
                return false;
            }
            startRow++;
            int MaxRow = sheet.LastRowNum;
            for (int y = startRow; y <= MaxRow; y++)
            {
                IRow row = sheet.GetRow(y);
                if (row == null)
                    continue;
                List<string> RowError = new List<string>();
                var value = row.GetCell(startCell + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                relatship.Add(value, y);
                foreach (var item in rules)
                {
                    if (!item.Rule.Check(row, startCell))
                    {
                        RowError.Add(item.Rule.Name);
                    }
                }
                if (RowError.Count() != 0)
                {
                    ErrorMessage.Add(value, RowError);
                }
            }
            return true;
        }


        /// <summary>
        /// 保存正确项目的数据
        /// </summary>
        /// <param name="filePath">提交表格</param>
        /// <param name="MasterPath">保存正确项目数据的文件路径</param>
        /// <param name="mistakes">错误信息</param>
        /// <param name="ErrorMessage">提交表格中存在填写错误的信息</param>
        /// <param name="relatship">提交表格的项目编号 与在该表格中行号</param>
        /// <returns></returns>
        public bool SaveCurrent(string filePath, string MasterPath, ref string mistakes, Dictionary<string, List<string>> ErrorMessage, Dictionary<string, int> relatship)
        {
            Dictionary<string, int> ship = new Dictionary<string, int>();
            int MasStartRow = 1;
            if (!GetProject(MasterPath, ref mistakes, ref ship, ref MasStartRow))
            {
                return false;
            }

            IWorkbook MasWorkbook = null;
            using (FileStream fs = new FileStream(MasterPath, FileMode.Open, FileAccess.Read))
            {
                MasWorkbook = WorkbookFactory.Create(fs);
            }
            if (MasWorkbook == null)
            {
                mistakes = "打开数据总表失败：" + MasterPath;
                return false;
            }
            ISheet MasSheet = MasWorkbook.GetSheetAt(0);


            IWorkbook workbook = null;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }
            if (workbook == null)
            {
                mistakes = "打开文件失败：" + filePath;
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            if (!FindHeader(sheet, ref startRow, ref startCell))
            {
                mistakes = "未找到表头：" + filePath;
                return false;
            }
            foreach (var item in relatship.Keys)
            {
                if (ErrorMessage.ContainsKey(item))
                    continue;
                if (ship.ContainsKey(item))
                    continue;
                IRow MasRow = MasSheet.GetRow(MasStartRow++);

                IRow row = sheet.GetRow(relatship[item]);

                for (int x = startCell, y = 0; x < row.LastCellNum; x++, y++)
                {
                    ICell cell = row.GetCell(x);
                    if (cell == null)
                        continue;

                    ICell MasCell = MasRow.GetCell(y);

                    switch (cell.CellType)
                    {
                        case CellType.Boolean:
                            MasCell.SetCellValue(cell.BooleanCellValue);
                            break;
                        case CellType.Numeric:
                            MasCell.SetCellValue(cell.NumericCellValue);
                            break;
                        case CellType.String:
                            MasCell.SetCellValue(cell.StringCellValue);
                            break;
                    }
                }
            }
            using (FileStream fs = new FileStream(MasterPath, FileMode.Open, FileAccess.Write))
            {
                MasWorkbook.Write(fs);
            }

            return true;
        }




        /*
         * 根据错误信息来生成错误表格
         * 检索表格：FilePath
         * 输出错误表格：resultPath
         * 错误信息：ErrorInformation
         * 生成表格  过程中错误的信息 mistakes
         */
        public bool OutputError(string FilePath, string resultPath, Dictionary<string, List<string>> ErrorInformation, ref string mistakes)
        {
            IWorkbook workbook;
            try
            {
                FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                mistakes = er.Message;
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            if (!FindHeader(sheet, ref startRow, ref startCell))
            {
                mistakes = "未找到表头";
                return false;
            }
            int i = startRow + 1;
            IRow row = sheet.GetRow(i);
            int MaxRow = sheet.LastRowNum;
            while (row != null)
            {
                var value = row.GetCell(startCell + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (ErrorInformation.ContainsKey(value))
                {
                    row = sheet.GetRow(++i);
                }
                else
                {
                    sheet.ShiftRows(i + 1, MaxRow, -1);
                    row = sheet.GetRow(i);
                }
            }

            try
            {
                FileStream fs = new FileStream(resultPath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                mistakes = er.Message;
                return false;
            }
            return true;
        }


        /*检查
         * summaryPath 该区域的总表
         * subPath 本次检查所提交的表格
         * statusPath 将提交表格中正确的更新到总表之后，保存一份总表的副本
         * summaryErrorExcel 检查更新之后，总表中还存在的错误
         * subErrorExcel 检查提交表格之后，还存在的错误表格
         * fault本次检查之后，可能发生的错误
         */
        public bool Check(string summaryPath, string subPath, string statusPath, string summaryErrorExcel, string subErrorExcel, out string fault)
        {
            Dictionary<string, int> summaryShip = new Dictionary<string, int>();
            string Mistakes = null;
            /*检查总表*/
            if (!CheckExcel(summaryPath, ref Mistakes, ref summaryError, ref summaryShip))
            {
                fault = Mistakes;
                return false;
            }
            /*检查提交表格*/
            Dictionary<string, int> subShip = new Dictionary<string, int>();
            if (!CheckExcel(subPath, ref Mistakes, ref subError, ref subShip))
            {
                fault = Mistakes;
                return false;
            }
            if (!OutputError(subPath, subErrorExcel, subError, ref Mistakes))
            {
                fault = Mistakes;
                return false;
            }
            IWorkbook summWorkbook;
            try
            {
                FileStream fs = new FileStream(summaryPath, FileMode.Open, FileAccess.Read);
                summWorkbook = WorkbookFactory.Create(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                Mistakes = er.Message;
                fault = Mistakes;
                return false;
            }
            IWorkbook subWorkbook;
            try
            {
                FileStream fs = new FileStream(subPath, FileMode.Open, FileAccess.Read);
                subWorkbook = WorkbookFactory.Create(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                Mistakes = er.Message;
                fault = Mistakes;
                return false;
            }
            ISheet summsheet = summWorkbook.GetSheetAt(0);
            ISheet subsheet = subWorkbook.GetSheetAt(0);
            int sumStartRow = 0, sumStartCell = 0;
            if (!FindHeader(summsheet, ref sumStartRow, ref sumStartCell))
            {
                Mistakes = "未找到表头";
                fault = Mistakes;
                return false;
            }
            int subStartRow = 0, subStartCell = 0;
            if (!FindHeader(subsheet, ref subStartRow, ref subStartCell))
            {
                Mistakes = "未找到提交表格表头";
                fault = Mistakes;
                return false;
            }
            /*遍历提交表格中的每一行
             * 根据检查出来的错误来判断该行正确与否
             * 当该行正确的时候，假如总表中错误 则更新该行到总表  
             */
            foreach (string item in subShip.Keys)
            {
                if (subError.ContainsKey(item))
                    continue;
                if (summaryError.ContainsKey(item))
                {
                    int summRowNumber = summaryShip[item];
                    int subRowNumber = subShip[item];

                    IRow summRow = summsheet.GetRow(summRowNumber);
                    IRow subRow = subsheet.GetRow(subRowNumber);

                    int subMaxCellNum = subRow.LastCellNum;
                    for (int x1 = subStartCell, x2 = sumStartCell; x1 <= subMaxCellNum; x1++, x2++)
                    {
                        ICell subCell = subRow.GetCell(x1);
                        if (subCell == null)
                            continue;
                        ICell sumCell = summRow.GetCell(x2);
                        if (sumCell == null)
                            sumCell = summRow.CreateCell(x2, subCell.CellType);
                        switch (subCell.CellType)
                        {
                            case CellType.Boolean:
                                sumCell.SetCellValue(subCell.BooleanCellValue);
                                break;
                            case CellType.Numeric:
                                sumCell.SetCellValue(subCell.NumericCellValue);
                                break;
                            case CellType.String:
                                sumCell.SetCellValue(subCell.StringCellValue);
                                break;
                        }

                    }
                    summaryError.Remove(item);
                }

            }
            /*更新总表*/
            try
            {
                FileStream fs = new FileStream(summaryPath, FileMode.Open, FileAccess.Write);
                summWorkbook.Write(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                Mistakes = er.Message;
                fault = Mistakes;
                return false;
            }
            /*保存总表副本*/
            try
            {
                FileStream fs = new FileStream(statusPath, FileMode.Create, FileAccess.Write);
                summWorkbook.Write(fs);
                fs.Close();
            }
            catch (Exception er)
            {
                fault = er.Message;
                return false;
            }
            /*保存总表中存在错误的行*/
            if (!OutputError(summaryPath, summaryErrorExcel, summaryError, ref Mistakes))
            {
                Mistakes = "保存总表错误失败；";
                fault = Mistakes;
                return false;
            }
            fault = Mistakes;
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