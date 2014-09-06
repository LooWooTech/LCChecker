using LCChecker.Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    public class BaseController : Controller
    {
        private LCDbContext db = new LCDbContext();

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


        //创建表格(拷贝 但是没有简单的拷贝 拷贝的那份整齐)  Path新创建表格的路径 OrigionPath原数据表格路径 将混乱（或者从0,0开始）的数据表格拷贝一份
        public bool CreateExcel(string Path, string OrigionPath,out string mistakes)
        {
            XSSFWorkbook NewWorkbook = new XSSFWorkbook();
            ISheet NewSheet = NewWorkbook.CreateSheet("表1");
            IWorkbook OldWorkbook;
            try
            {
                FileStream os = new FileStream(OrigionPath, FileMode.Open, FileAccess.Read);
                OldWorkbook = WorkbookFactory.Create(os);
            }
            catch (Exception er)
            {
                mistakes = er.Message;
                return false;
            }
            ISheet OldSheet = OldWorkbook.GetSheetAt(0);
            int OldRow = 0, OldCell = 0;
            if (!FindHeader(OldSheet, ref OldRow, ref OldCell))
            {
                mistakes = "未能找到表格：" + OrigionPath + "的表头，导致未能检索表格数据";
                return false;
            }
            int h = 0;
            for (int i = OldRow; i <= OldSheet.LastRowNum; i++)
            {
                IRow orow = OldSheet.GetRow(i);
                IRow NewRow = NewSheet.CreateRow(h++);
                for (int j = OldCell, l = 0; j < orow.LastCellNum; j++, l++)
                {
                    var oCell = orow.GetCell(j);
                    if (oCell == null)
                        continue;
                    var nCell = NewRow.CreateCell(l, oCell.CellType);
                    switch (oCell.CellType)
                    {
                        case CellType.Boolean:
                            nCell.SetCellValue(oCell.BooleanCellValue);
                            break;
                        case CellType.Numeric:
                            nCell.SetCellValue(oCell.NumericCellValue);
                            break;
                        case CellType.String:
                            nCell.SetCellValue(oCell.StringCellValue);
                            break;
                    }
                }
            }
            try
            {
                FileStream fs = new FileStream(Path, FileMode.Create);
                NewWorkbook.Write(fs);
            }
            catch(Exception er)
            {
                mistakes = er.Message;
                return false;
            }
            mistakes = "";
            return true;

        }

        //根据错误行 将存在错误的表格中相应的错误行 独立保存一份错误表格 Path即要创建的错误表格路径 origPath为存在错误的表格路径 error错误信息
        public bool  NewExcel(string Path, string origPath, Dictionary<string, List<string>> error,out string mistakes)
        {
            //新建表
            XSSFWorkbook eWorkbook = new XSSFWorkbook();
            ISheet eSheet = eWorkbook.CreateSheet("错误行");
            IRow rHeader = eSheet.CreateRow(0);
            for (int i = 0; i < 43; i++)
            {
                ICell cell = rHeader.CreateCell(i);
                //cell.CellStyle = GetCellStyle(eWorkbook, stylexls.小头);
                cell.SetCellValue(string.Format("{0}栏",i+1));
            }
            int rNumber = 1;
            IWorkbook oWorkbook;
            try
            {
                FileStream os = new FileStream(origPath, FileMode.Open, FileAccess.ReadWrite);
                //数据表
                oWorkbook = WorkbookFactory.Create(os);
            }
            catch (Exception er)
            {
                mistakes = er.Message;
                return false;
            }
            ISheet oSheet = oWorkbook.GetSheetAt(0);
            int startRow = 0, startCell = 0;
            //找到原始数据开始的行 列
            if (!FindHeader(oSheet, ref startRow, ref startCell))
            {
                mistakes = "未能找到源数据表格,导致未能检索表格数据";
                return false;
            }
            startRow++;
            for (int i = startRow; i <= oSheet.LastRowNum; i++)
            {
                IRow oRow = oSheet.GetRow(i);
                var value = oRow.GetCell(startCell + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (error.ContainsKey(value))
                {
                    IRow NRow = eSheet.CreateRow(rNumber++);
                    for (int j = startCell, CellNumber = 0; j <= startCell + 43; j++, CellNumber++)
                    {
                        var oCell = oRow.GetCell(j);
                        if (oCell == null)
                            continue;
                        var nCell = NRow.CreateCell(CellNumber, oCell.CellType);

                       // nCell.CellStyle = GetCellStyle(eWorkbook,stylexls.默认);
                        
                        switch (oCell.CellType)
                        {
                            case CellType.Boolean:
                                nCell.SetCellValue(oCell.BooleanCellValue);
                                break;
                            case CellType.Numeric:
                                nCell.SetCellValue(oCell.NumericCellValue);
                                break;
                            case CellType.String:
                                nCell.SetCellValue(oCell.StringCellValue);
                                break;
                        }
                    }
                }
            }
            try
            {
                FileStream fs = new FileStream(Path, FileMode.Create);
                //FileStream fs = System.IO.File.OpenWrite(Path);
                eWorkbook.Write(fs);
            }
            catch (Exception er){
                mistakes = er.Message;
                return false;
            }
            mistakes = "";
            return true;
        }


        



        

        /*
         * 数据采集  
         * 采集数据种类：项目个数 flag 用于判断是否第一次采集 /更新数据
         */
        public bool CollectData(string region)
        {
            string errorInformation = null;
            string DataPath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region), "summary.xlsx");
            //检查该表中存在错误的行
            DetectEngine Engine = new DetectEngine(DataPath);
            Engine.CheckSummaryExcel(DataPath,out errorInformation);

            try
            {
                Detect record = db.DETECT.Where(x => x.region == region).FirstOrDefault();
                if (record == null)
                {
                    return false;
                }
                FileStream fs = new FileStream(DataPath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                int startRow = 0, startCell = 0;
                if (!FindHeader(sheet, ref startRow, ref startCell))
                {
                    return false;
                }
                int sum = sheet.LastRowNum - startRow;    
                record.sum = sum;
                record.Correct = sum - Engine.Error.Count();//正确的行（项目）个数 
                if (ModelState.IsValid)
                {
                    db.Entry(record).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }





        public enum stylexls
        {
            大头,
            小头,
            小小头,
            默认
        }

        public static ICellStyle GetCellStyle(IWorkbook workbook, stylexls str)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            IFont fontBigheader = workbook.CreateFont();
            fontBigheader.FontHeightInPoints = 22;
            fontBigheader.FontName = "微软雅黑";
            fontBigheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            IFont fontSmallheader = workbook.CreateFont();
            fontSmallheader.FontHeightInPoints = 14;
            fontSmallheader.FontName = "黑体";

            IFont fontText = workbook.CreateFont();
            fontText.FontHeightInPoints = 12;
            fontText.FontName = "宋体";

            IFont fontthinheader = workbook.CreateFont();
            fontthinheader.FontName = "宋体";
            fontthinheader.FontHeightInPoints = 11;
            fontthinheader.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;


            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            //边框颜色
            cellStyle.BottomBorderColor = HSSFColor.OliveGreen.Black.Index;
            cellStyle.TopBorderColor = HSSFColor.OliveGreen.Black.Index;

            //背景图形
            cellStyle.FillForegroundColor = HSSFColor.White.Index;
            cellStyle.FillBackgroundColor = HSSFColor.Black.Index;

            //文本对齐  左对齐  居中  右对齐  现在是居中
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            //垂直对齐
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            //自动换行
            cellStyle.WrapText = true;

            //缩进
            cellStyle.Indention = 0;

            switch (str)
            {
                case stylexls.大头:
                    cellStyle.SetFont(fontBigheader);
                    break;
                case stylexls.小头:
                    cellStyle.SetFont(fontSmallheader);
                    break;
                case stylexls.默认:
                    cellStyle.SetFont(fontText);
                    break;
                case stylexls.小小头:
                    cellStyle.SetFont(fontthinheader);
                    break;
            }


            return cellStyle;
        }


        public bool AddCssExcel(string filePath,out string mistakes,out MemoryStream stream)
        {
            IWorkbook workbook;
            stream = new MemoryStream();
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);
            }
            catch (Exception er)
            {
                mistakes = er.Message;
                stream = null;
                //workbook = null;
                return false;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            row.Height = 20 * 20;
            for (int x = 0; x < 43; x++)
            {
                ICell cell = row.GetCell(x);
                cell.CellStyle = GetCellStyle(workbook, stylexls.小小头);
            }

            for (int y = 1; y < sheet.LastRowNum; y++)
            {
                IRow drow = sheet.GetRow(y);
                for (int x = 0; x < 43; x++)
                {
                    ICell cell = drow.GetCell(x);
                    cell.CellStyle = GetCellStyle(workbook, stylexls.默认);
                }
            }
            mistakes = "";
            stream = null;
            workbook.Write(stream);
            stream.Flush();
            stream.Position = 0;
            return true;
        }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        //List<RuleInfo> rules = new List<RuleInfo>();
        //public void initRule()
        //{
        //    var list = new List<IRowRule>();

        //    list.Add(new NoLessThanRowRule() { Column1Index = 5, Column2Index = 6 });
        //    list.Add(new SumRowRule() { SumColumnIndex = 6, ColumnIndices = new[] { 18, 23 } });
        //    list.Add(new SumRowRule() { SumColumnIndex = 13, ColumnIndices = new[] { 14, 15 } });
        //    list.Add(new SumRowRule() { SumColumnIndex = 18, ColumnIndices = new[] { 19, 20 } });
        //    list.Add(new SumRowRule() { SumColumnIndex = 20, ColumnIndices = new[] { 21, 22 } });
        //    list.Add(new SumRowRule() { SumColumnIndex = 23, ColumnIndices = new[] { 24, 27, 28, 31, 32 } });
        //    list.Add(new SumRowRule() { SumColumnIndex = 24, ColumnIndices = new[] { 25, 26 } });
        //    list.Add(new SumRowRule() { SumColumnIndex = 28, ColumnIndices = new[] { 29, 30 } });
        //    list.Add(new SumRowRule() { SumColumnIndex = 32, ColumnIndices = new[] { 33, 34 } });

        //    list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" } });
        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = new CellEmptyRowRule() { ColumnIndex = 17, isEmpty = false, isNumeric = false },
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 18, isEmpty = true, isNumeric = true }
        //    });
        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = new CellEmptyRowRule() { ColumnIndex = 17, isEmpty = true, isNumeric = false },
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 18, isEmpty = false, isNumeric = true }
        //    });

        //    list.Add(new UniqueValueRowRule() { ColumnIndex = 3, Keyword = "综合整治" });

        //    var rule1 = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是" } };

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 8, isEmpty = true, isNumeric = false }
        //    });
        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 9, isEmpty = true, isNumeric = false }
        //    });
        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 10, isEmpty = true, isNumeric = false }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 11, isEmpty = true, isNumeric = false }
        //    });
        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 12, isEmpty = true, isNumeric = false }
        //    });
        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 13, isEmpty = true, isNumeric = true }
        //    });
        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 16, isEmpty = true, isNumeric = false }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 19, isEmpty = true, isNumeric = true }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 27, isEmpty = true, isNumeric = false }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 28, isEmpty = true, isNumeric = true }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule1,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 31, isEmpty = true, isNumeric = false }
        //    });

        //    var rule2 = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "否" } };

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule2,
        //        Rule =
        //            new CellRangeRowRule()
        //            {
        //                ColumnIndex = 8,
        //                Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误", "3、属于复垦、整理、综合整治等项目无法确认信息项目" }
        //            }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule2,
        //        Rule = new CellEmptyRowRule() { ColumnIndex = 31, isEmpty = true, isNumeric = false }
        //    });


        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule2,
        //        Rule = new ConditionalRowRule()
        //        {
        //            Condition = new CellRangeRowRule()
        //            {
        //                ColumnIndex = 8,
        //                Values = new[] { "1、调剂出项目对方指标使用有误", "2、本县自行补充(含尚未调剂出)项目有误" }
        //            },
        //            Rule = new MultipleCellRangeRowRule()
        //            {
        //                ColumnIndices = new[] { 10, 11, 12, 16, 19, 27, 31 },
        //                isAny = true,
        //                isEmpty = false,
        //                isNumeric = false
        //            }
        //        }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule2,
        //        Rule = new ConditionalRowRule()
        //        {
        //            Condition = new CellRangeRowRule()
        //            {
        //                ColumnIndex = 8,
        //                Values = new[] { "1、调剂出项目对方指标使用有误" }
        //            },
        //            Rule = new CellEmptyRowRule()
        //            {
        //                ColumnIndex = 19,
        //                isEmpty = false,
        //                isNumeric = true
        //            }
        //        }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule2,
        //        Rule = new ConditionalRowRule()
        //        {
        //            Condition = new CellRangeRowRule()
        //            {
        //                ColumnIndex = 8,
        //                Values = new[] { "2、本县自行补充(含尚未调剂出)项目有误" }
        //            },
        //            Rule = new MultipleCellRangeRowRule()
        //            {
        //                ColumnIndices = new[] { 10, 11, 12, 16, 27, 31 },
        //                isAny = false,
        //                isEmpty = false,
        //                isNumeric = true
        //            }
        //        }
        //    });

        //    list.Add(new ConditionalRowRule()
        //    {
        //        Condition = rule2,
        //        Rule = new ConditionalRowRule()
        //        {
        //            Condition = new CellRangeRowRule()
        //            {
        //                ColumnIndex = 8,
        //                Values = new[] { "3、属于复垦、整理、综合整治等项目无法确认信息项目" }
        //            },
        //            Rule = new AndRule()
        //            {
        //                Rule1 = new MultipleCellRangeRowRule()
        //                {
        //                    ColumnIndices = new[] { 27, 28 },
        //                    isAny = false,
        //                    isEmpty = false,
        //                    isNumeric = true
        //                },
        //                Rule2 = new MultipleCellRangeRowRule()
        //                {
        //                    ColumnIndices = new[] { 23, 32 },
        //                    isAny = false,
        //                    isEmpty = true,
        //                    isNumeric = true
        //                }
        //            }
        //        }
        //    });


        //    foreach (var item in list)
        //    {
        //        rules.Add(new RuleInfo() { Rule=item});
        //    }
        //}
        //public void CheckExcel()
        //{
        //    initRule();
 
        //}
        ///*检查新上传的表格
        // * 检查每一行存在的问题 
        // * filePath表示为需要检查的表格路径 summaryPath总表路径 region为相应的区域  submit提交的次数
        // * 过程：1、新建错误表格（填充表头），在检查的时候 将错误行保存到错误表格
        // */
        //public List<Mistake> CheckExcel(string filePath,string summaryPath,string region,int submit)
        //{
        //    initRule();
        //    List<Mistake> Result = new List<Mistake>();
        //    int startRow = 0, startCell = 0;//查找上传的数据表头
        //    using (FileStream fs = new FileStream(filePath,FileMode.Open,FileAccess.ReadWrite))//打开检查的表格
        //    {
        //        using (FileStream sums = new FileStream(summaryPath, FileMode.Open, FileAccess.ReadWrite))//打开需要更新数据的表格
        //        {
        //            /*新建错误表格*/
        //            string ErrorPath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region), submit + "错误" + ".xls");
        //            FileStream es = new FileStream(ErrorPath, FileMode.Create, FileAccess.Write);
        //            HSSFWorkbook errorbook = new HSSFWorkbook(es);
        //            ISheet Ersheet = errorbook.CreateSheet();
        //            IRow errorRow = Ersheet.CreateRow(0);
        //            for (int i = 0; i < 43; i++)
        //            {
        //                errorRow.CreateCell(i).SetCellValue(string.Format("{0}栏", i + 1));
        //            }


        //            /*NPOI打开总表*/
        //            XSSFWorkbook summarybook = new XSSFWorkbook(sums);
        //            ISheet summarySheet = summarybook.GetSheetAt(0);
        //            int summaryRow = 0, summaryCell = 0;
        //            FindHeader(summarySheet, ref summaryRow, ref summaryCell);


        //            /*NPOI打开刚上传的表格*/
        //            XSSFWorkbook workbook = new XSSFWorkbook(fs);
        //            ISheet sheet = workbook.GetSheetAt(0);
        //            if (!FindHeader(sheet, ref startRow, ref startCell))
        //            {
        //                Result.Add(new Mistake() { Error = "没有找到表头", row = 0 });
        //                return Result;
        //            }
        //            startRow++;//因为表头占据一行 所以要从startRow行开始检查数据
        //            var Row = sheet.GetRow(startRow);
        //            bool flag = false;
        //            int ERowNumber = 1;//错误表格填写从第2行开始
        //            while (Row != null)
        //            {
        //                flag = false;
        //                foreach (var item in rules)
        //                {
        //                    if (!item.Rule.Check(Row, startCell))
        //                    {
        //                        Result.Add(new Mistake() { Error = item.Rule.Name, row = startRow });
        //                        flag = true;
        //                    }
        //                }
        //                if (flag)//该行存在错误，错误表新建一行 往这一行中填充数据
        //                {
        //                    IRow erow = Ersheet.CreateRow(ERowNumber);
        //                    int m = startCell;
        //                    for (int i = 0; i < 43; i++)
        //                    {
        //                        erow.CreateCell(i).SetCellValue(Row.GetCell(m++, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim());
        //                    }
        //                }
        //                    //数据正确
        //                else {
        //                    var value2 = Row.GetCell(startCell, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();//获取当前行的首列（数据开始：序号）
        //                    for (int i = summaryRow+1; i < summarySheet.LastRowNum; i++)
        //                    {
        //                        var value1 = summarySheet.GetRow(i).GetCell(summaryCell, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();     
        //                        if (value1 == value2)//找到总表中相应的数据行 第i行
        //                        {
        //                            for (int j = summaryCell,k=startCell; j < summaryCell + 43; j++,k++)//从第i行第summaryCell列开始的地方更新数据 一共获取43列数据
        //                            {
        //                                value1 = Row.GetCell(k, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();//从当前行的第startCell列获取数据
        //                                summarySheet.CreateRow(i).CreateCell(j).SetCellValue(value1);
        //                            }
        //                        }
        //                    }
        //                }
        //                startRow++;
        //                Row = sheet.GetRow(startRow);
        //            }
        //        }
                


                
                
        //    }
        //    return Result;
        //}
        ///*复制表格 已成功 提交次数 区域*/
        //public void CopyExcel(int submit,string region)
        //{
        //    if (submit == 1)
        //    {
        //        return;
        //    }
        //    string origionalPath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + (submit - 1).ToString()), region + ".xls");
        //    string NowPath = Path.Combine(HttpContext.Server.MapPath("../Uploads/" + region + "/" + submit), region + ".xls");
        //    System.IO.File.Copy(origionalPath, NowPath);
        //}

        
    }
}
