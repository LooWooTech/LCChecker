using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LooWoo.Land.LandControlChecker.Rules;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace LooWoo.Land.LandControlChecker
{
    public class CheckStatusChangedEventArgs: EventArgs
    {
        public string File { get; set; }

        public string Text { get; set; }
    }

    public class CheckEngine
    {
        private readonly IList<RuleInfo> rules;

        public event EventHandler<CheckStatusChangedEventArgs> CheckStatusChanged;

        public CheckEngine()
        {
            var list = new List<IRowRule>();

            list.Add(new NoLessThanRowRule() {Column1Index = 5, Column2Index = 6});
            list.Add(new SumRowRule() {SumColumnIndex = 6, ColumnIndices = new[] {18, 23}});
            list.Add(new SumRowRule() {SumColumnIndex = 6, ColumnIndices = new[] {13, 18, 23}});
            list.Add(new SumRowRule() {SumColumnIndex = 13, ColumnIndices = new[] {14, 16}});
            list.Add(new SumRowRule() {SumColumnIndex = 18, ColumnIndices = new[] {19, 20}});
            list.Add(new SumRowRule() {SumColumnIndex = 20, ColumnIndices = new[] {21, 22}});
            list.Add(new SumRowRule() {SumColumnIndex = 23, ColumnIndices = new[] {24, 27, 28, 31, 32}});
            list.Add(new SumRowRule() {SumColumnIndex = 24, ColumnIndices = new[] {25, 27}});
            list.Add(new SumRowRule() {SumColumnIndex = 28, ColumnIndices = new[] {29, 30}});
            list.Add(new SumRowRule() {SumColumnIndex = 32, ColumnIndices = new[] {33, 35}});

            list.Add(new CellRangeRowRule() {ColumnIndex = 7, Values = new[] {"是", "否"}});
            list.Add(new ConditionalRowRule()
                {
                    Condition = new CellEmptyRowRule() {ColumnIndex = 17, isEmpty = false, isNumeric = false},
                    Rule = new CellEmptyRowRule() {ColumnIndex = 18, isEmpty = true, isNumeric = true}
                });
            list.Add(new ConditionalRowRule()
                {
                    Condition = new CellEmptyRowRule() {ColumnIndex = 17, isEmpty = true, isNumeric = false},
                    Rule = new CellEmptyRowRule() {ColumnIndex = 18, isEmpty = false, isNumeric = true}
                });
            list.Add(new UniqueValueRowRule() {ColumnIndex = 3, Keyword = "综合整治"});

            var rule1 = new CellRangeRowRule() {ColumnIndex = 7, Values = new[] {"是"}};

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 8, isEmpty = true, isNumeric = false}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 9, isEmpty = true, isNumeric = false}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 10, isEmpty = true, isNumeric = false}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 11, isEmpty = true, isNumeric = false}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 12, isEmpty = true, isNumeric = false}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 13, isEmpty = true, isNumeric = true}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 16, isEmpty = true, isNumeric = false}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 19, isEmpty = true, isNumeric = true}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 27, isEmpty = true, isNumeric = false}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 28, isEmpty = true, isNumeric = true}
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 31, isEmpty = true, isNumeric = false}
                });

            var rule2 = new CellRangeRowRule() {ColumnIndex = 7, Values = new[] {"否"}};

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule2,
                    Rule =
                        new CellRangeRowRule()
                            {
                                ColumnIndex = 8,
                                Values = new[] {"1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误", "3、属于复垦、整理、综合整治等项目无法确认信息项目"}
                            }
                });

            list.Add(new ConditionalRowRule()
                {
                    Condition = rule2,
                    Rule = new CellEmptyRowRule() {ColumnIndex = 31, isEmpty = true, isNumeric = false}
                });


            list.Add(new ConditionalRowRule()
                {
                    Condition = rule2,
                    Rule = new ConditionalRowRule()
                        {
                            Condition = new CellRangeRowRule()
                                {
                                    ColumnIndex = 8,
                                    Values = new[] {"1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误"}
                                },
                            Rule = new MultipleCellRangeRowRule()
                                {
                                    ColumnIndices = new[] {10, 11, 12, 16, 19, 27, 31},
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
                                    Values = new[] {"1、调剂出项目对方指标使用有误"}
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
                                    Values = new[] {"2、本县自行补充(含尚未调剂出)项目有误"}
                                },
                            Rule = new MultipleCellRangeRowRule()
                                {
                                    ColumnIndices = new[] {10, 11, 12, 16, 27, 31},
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
                                    Values = new[] {"3、属于复垦、整理、综合整治等项目无法确认信息项目"}
                                },
                            Rule = new AndRule()
                                {
                                    Rule1 = new MultipleCellRangeRowRule()
                                        {
                                            ColumnIndices = new[] {27, 28},
                                            isAny = false,
                                            isEmpty = false,
                                            isNumeric = true
                                        },
                                    Rule2 = new MultipleCellRangeRowRule()
                                        {
                                            ColumnIndices = new[] {23, 32},
                                            isAny = false,
                                            isEmpty = true,
                                            isNumeric = true
                                        }
                                }
                        }
                });


            var list2 = new List<RuleInfo>();
            var col = 12;
            var index = 0;
            for (int i = 0; i < 10; i++)
            {
                list2.Add(new RuleInfo()
                    {
                        Rule = list[index],
                        CheckSheetColumnIndex = col + i,
                        SheetIndex = 1 + i
                    });
                index ++;

            }

            list2.Add(new RuleInfo()
                {
                    Rule = list[index],
                    CheckSheetColumnIndex = 22,
                    SheetIndex = -1
                });

            index++;

            list2.Add(new RuleInfo()
                {
                    Rule = list[index],
                    CheckSheetColumnIndex = 23,
                    SheetIndex = 11
                });

            index++;

            list2.Add(new RuleInfo()
                {
                    Rule = list[index],
                    CheckSheetColumnIndex = 24,
                    SheetIndex = 12
                });
            index++;

            list2.Add(new RuleInfo()
                {
                    Rule = list[index],
                    CheckSheetColumnIndex = 25,
                    SheetIndex = 13
                });
            index++;

            col = 27;
            for (int i = 0; i < 11; i++)
            {
                list2.Add(new RuleInfo()
                    {
                        Rule = list[index],
                        CheckSheetColumnIndex = col + i,
                        SheetIndex = 14 + i
                    });
                index ++;

            }

            col = 40;
            for (int i = 0; i < 6; i++)
            {
                list2.Add(new RuleInfo()
                    {
                        Rule = list[index],
                        CheckSheetColumnIndex = col + i,
                        SheetIndex = 25 + i
                    });
                index ++;
            }

            rules = list2;
        }

        private void RowCopy(IRow from, int toRowIndex, ISheet sheet, int count)
        {
            var row = sheet.GetRow(toRowIndex);
            if (row == null)
                row = sheet.CreateRow(toRowIndex);

            for (int i = 0; i < count; i++)
            {
                var fromCell = from.GetCell(i);
                if (fromCell == null) continue;

                var cell = row.GetCell(i);
                if (cell == null)
                    cell = row.CreateCell(i, fromCell.CellType);

                //cell.CellStyle = fromCell.CellStyle;
                //if (fromCell.CellFormula == "")
                switch (fromCell.CellType)
                {
                    case CellType.Boolean:
                        cell.SetCellValue(fromCell.BooleanCellValue);
                        break;
                    case CellType.Numeric:
                        cell.SetCellValue(fromCell.NumericCellValue);
                        break;
                    case CellType.String:
                        cell.SetCellValue(fromCell.StringCellValue);
                        break;
                }
                   
            }
        }

        public int Check(string inputDir, string resultfile, bool cityLevel)
        {
            DirectoryInfo di = new DirectoryInfo(inputDir);
            var files = new List<FileInfo>(di.GetFiles("*.xls"));
            //files.AddRange(di.GetFiles("*.xlsx"));

            var resultStartRowDict = new Dictionary<int, int>();
            foreach (var item in rules)
            {
                resultStartRowDict.Add(item.Id, 1);
            }

            var result = new Dictionary<string, Dictionary<int, int>>();
            var count = 0;
            foreach (FileInfo file in files)
            {
                CheckStatusChanged(this, new CheckStatusChangedEventArgs()
                    {
                        File = Path.GetFileName(file.FullName),
                        Text = "开始检查..."
                    });
                Dictionary<string, Dictionary<int, int>> result2;
                var ret = Check(file.FullName, resultfile, resultStartRowDict, cityLevel, out result2);
                if (ret) count++;
                Merge(result2, result);
            }

            OutputSummary(result, resultfile);
            return count;
        }

        private void OutputSummary(Dictionary<string, Dictionary<int, int>> summary, string resultfile)
        {
            IWorkbook wb;

            using (var stream = new FileStream(resultfile, FileMode.Open, FileAccess.ReadWrite))
            {
                wb = WorkbookFactory.Create(stream);
                ISheet sheet = wb.GetSheetAt(0);

                var startRow = 5;
                var signalCol = 0;
                var regionCol = 1;

                var rowIndex = startRow;
                while (true)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null ||
                        string.IsNullOrEmpty(
                            row.GetCell(signalCol, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim()))
                    {
                        break;
                    }

                    var region = row.GetCell(regionCol, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                    if (string.IsNullOrEmpty(region))
                    {
                        rowIndex++;
                        continue;
                    }

                    foreach (var pair in summary)
                    {
                        if (pair.Key.Contains(region))
                        {
                            foreach (var rule in rules)
                            {
                                if (pair.Value.ContainsKey(rule.Id))
                                {
                                    var cell = row.GetCell(rule.CheckSheetColumnIndex);
                                    if (cell == null) row.CreateCell(rule.CheckSheetColumnIndex);
                                    cell.SetCellValue(pair.Value[rule.Id].ToString());
                                }
                            }
                            break;
                        }
                    }

                    rowIndex++;
                }
            }
            using (var stream = new FileStream(resultfile, FileMode.Open, FileAccess.ReadWrite))
            {
                wb.Write(stream);
            }
        }


        private void Merge(Dictionary<string, Dictionary<int, int>> from, Dictionary<string, Dictionary<int, int>> to)
        {
            foreach (var dict in from)
            {
                if (to.ContainsKey(dict.Key))
                {
                    var dict2 = to[dict.Key];
                    foreach (var pair in dict.Value)
                    {
                        if (dict2.ContainsKey(pair.Key))
                        {
                            dict2[pair.Key] = dict2[pair.Key] + pair.Value;
                        }
                        else
                        {
                            dict2.Add(pair.Key, pair.Value);
                        }
                    }
                }
                else
                {
                    to.Add(dict.Key, dict.Value);
                }
            }
        }

        private bool Check(string inputfile, string resultfile, Dictionary<int, int> resultStartRowDict, bool cityLevel, out Dictionary<string, Dictionary<int, int>> summary)
        {
            IWorkbook wb, wb2;
            using (var stream = new FileStream(inputfile, FileMode.Open, FileAccess.ReadWrite))
            {
                wb = WorkbookFactory.Create(stream);
            }
            
            using (var stream2 = new FileStream(resultfile, FileMode.Open, FileAccess.ReadWrite))
            {
                wb2 = WorkbookFactory.Create(stream2);
            }
                
            ISheet sheet = wb.GetSheetAt(0);
            var result = new Dictionary<string, Dictionary<int, int>>();
            var rowIndex = 0;
            var colIndex = 0;
            if (FindHeader(sheet, ref rowIndex, ref colIndex) == false)
            {
                CheckStatusChanged(this, new CheckStatusChangedEventArgs()
                    {
                        File = Path.GetFileName(inputfile),
                        Text = "错误：无法找到表头"
                    });
                summary = result;
                return false;
            }
            rowIndex++;
            var row = sheet.GetRow(rowIndex);

            var rowCount = 0;
            var errorCount = 0;

            while (row != null)
            {
                var region = row.GetCell(colIndex+1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(region))
                    break;

                var tokens = region.Split(',');
                if (tokens.Length == 3)
                {
                    rowCount++;

                    var key = (cityLevel ? tokens[1] : (tokens[1] + "-" + tokens[2]));

                    if (result.ContainsKey(key) == false)
                    {
                        var dict = new Dictionary<int, int>();
                        foreach (var rule in rules)
                        {
                            if(rule.Enabled) dict.Add(rule.Id, 0);
                        }
                        result.Add(key, dict);
                    }

                    var d = result[key];
                    foreach (var rule in rules)
                    {
                        if (rule.Enabled && rule.Rule.Check(row, colIndex) == false)
                        {
                            d[rule.Id] = d[rule.Id]+1;
                            var sheet2 = wb2.GetSheetAt(rule.SheetIndex);
                            RowCopy(row, resultStartRowDict[rule.Id], sheet2, 34);
                            resultStartRowDict[rule.Id] = resultStartRowDict[rule.Id] + 1;
                            errorCount++;
                        }
                    }

                    if (rowCount%20 == 0)
                    {
                        CheckStatusChanged(this, new CheckStatusChangedEventArgs()
                            {
                                File = Path.GetFileName(inputfile),
                                Text = string.Format("已扫描 {0} 行，发现 {1} 个错误", rowCount, errorCount)
                            });
                    }
                }
                rowIndex++;
                
                row = sheet.GetRow(rowIndex);
            }

            if (rowCount%20 != 0)
            {
                CheckStatusChanged(this, new CheckStatusChangedEventArgs()
                    {
                        File = Path.GetFileName(inputfile),
                        Text = string.Format("已扫描 {0} 行，发现 {1} 个错误", rowCount, errorCount)
                    });
            }

            using (var stream2 = new FileStream(resultfile, FileMode.Open, FileAccess.ReadWrite))
            {
                wb2.Write(stream2);
            };
            summary = result;
            return true;
        }

        private bool FindHeader(ISheet sheet, ref int startrow, ref int startcol)
        {
            for (var i = 0; i < 5; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    for (var j = 0; j < 5; j++)
                    {
                        var value = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (value == "1栏")
                        {
                            for (var k = 1; k < 43; k++)
                            {
                                value = row.GetCell(k+j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                                if (value != string.Format("{0}栏", k+1))
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
