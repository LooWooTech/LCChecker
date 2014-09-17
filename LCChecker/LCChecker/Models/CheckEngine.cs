using LCChecker.Rules;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace LCChecker.Models
{
    public class CheckEngine
    {
        public List<RuleInfo> rules = new List<RuleInfo>();

        public Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();

        public Dictionary<string, Index2> Ship = new Dictionary<string, Index2>();

        public Dictionary<string, List<string>> GetError() {
            return Error;
        }
        public bool Check(string filePath, ref string mistakes)
        {
            int startRow = 0, startCell = 0;
            ISheet sheet = OpenSheet(filePath, true, ref startRow, ref startCell,  ref mistakes);
            if (sheet == null)
            {
                Error.Add("表格格式内容", new List<string> { "提交的表格无法检索 请核对格式" });
                return false;
            }

            startRow++;
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    break;
                List<string> ErrorRow = new List<string>();
                var value = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                foreach (var item in rules)
                {
                    if (!item.Rule.Check(row, startCell))
                    {
                        ErrorRow.Add(item.Rule.Name);
                    }
                }
                if (ErrorRow.Count() != 0)
                {
                    Error.Add(value, ErrorRow);
                }
            }

            return true;
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

        //public IWorkbook GetWorkbook(string filePath, ref string Mistakes)
        //{
        //    IWorkbook workbook = null;
        //    try
        //    {
        //        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        //        {
        //            workbook = WorkbookFactory.Create(fs);
        //        }
        //    }
        //    catch
        //    {
        //        Mistakes = "打开文件失败";
        //        return null;
        //    }
        //    if (workbook == null)
        //    {
        //        Mistakes = "打开文件失败";
        //        return null;
        //    }
        //    return workbook;
        //}


        /// <summary>
        /// 打开Excel，找第一个Sheet，并返回表头
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="findHeader"></param>
        /// <returns></returns>
        public ISheet OpenSheet(string filePath, bool findHeader, ref int startRow, ref int startCol, ref string errMsg)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch
            {
                errMsg = "打开Excel表格失败：";// +filePath;
                return null;
            }

            if (workbook == null)
            {
                errMsg = "打开Excel表格失败：";// +filePath;
                return null;
            }

            if (workbook.NumberOfSheets == 0)
            {
                errMsg = "Excel文件中没有表格。";
                return null;
            }

            var sheet = workbook.GetSheetAt(0);

            if (findHeader == false) return sheet;

            if (FindHeader(sheet, ref startRow, ref startCol) == false)
            {
                errMsg = "未找到文件表头[编号 市 县]";
                return null;
            }
            return sheet;
        }

        /// <summary>
        /// 获取全部项目数据
        /// </summary>
        /// <param name="filePath">正确项目文件</param>
        public void GetMessage(string filePath)
        {
            string fault = "";
            int startRow = 1,starCell=0;
            ISheet sheet = OpenSheet(filePath, false,ref startRow,ref starCell, ref fault);
            if (sheet == null)
            {
                Error.Add("表格内容格式", new List<string>() { "获取项目总表失败" });
                return;
            }
            
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                //项目编号
                var value = row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    break;
                //项目名称
                var value1 = row.Cells[3].StringCellValue;
                //项目地区
                var cityName = row.Cells[1].StringCellValue.Split(',');
                //新增耕地面积
                var AddArea = row.GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                double data = double.Parse(AddArea);
                //等别
                var grade = row.GetCell(35, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();

                //表8  14栏
                var Indicator = row.GetCell(13, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                double data2 = double.Parse(Indicator);

                if (cityName.Length < 3)
                {
                    continue;
                }
                Ship.Add(value, new Index2
                {
                    City = cityName[1],//市
                    County = cityName[2],//县
                    Name = value1,//项目名称
                    AddArea=data,//新增耕地
                    Indicators=data2,//  表 8   14栏
                    Grade=grade//等别
                });
            }
        }

        /// <summary>
        /// 获取补充耕地项目中的数据
        /// </summary>
        /// <param name="filePath"></param>
        public void GetSuppleMessage(string filePath)
        {
            string fault = "";
            int startRow = 1,startCell=0;
            ISheet sheet = OpenSheet(filePath, false, ref startRow, ref startCell, ref fault);
            if (sheet == null)
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
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                if (!engine.Rule.Check(row))
                    continue;
                //项目编号
                var value = row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    break;
                //项目名称
                var value1 = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                //市  县
                var cityName = row.Cells[1].StringCellValue.Split(',');
                //新增耕地面积
                var AddArea = row.GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                double data = double.Parse(AddArea);
                //等别
                var grade = row.GetCell(35, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                
                //表8  14栏
                var Indicator = row.GetCell(13,MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                double data2 = double.Parse(Indicator);

                var val = row.GetCell(36, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();

                Land LandData = GetLand(val);

                if (cityName.Length < 3)
                {
                    continue;
                }
                Ship.Add(value, new Index2
                {
                    City = cityName[1],//市
                    County = cityName[2],//县
                    Name = value1,//名称
                    AddArea = data,//新增耕地面积
                    Grade = grade,//质量等别
                    Indicators=data2,//表8中的14栏
                    Land=LandData//水田  水浇地 旱地
                });
            }
        }



        public Land GetLand(string value)
        {
            Land land = new Land();
            string val = value.Replace("，", ",").Replace(".", string.Empty).Replace("。", string.Empty);
            var team = val.Split(',');
            foreach (string item in team)
            {
                Regex r = new Regex(@"-?[0-9]");
                string data = r.Match(item).ToString();
                int position = item.IndexOf(data);
                if (position == 0)
                    continue;
                string ground = item.Substring(0,position);
                string area = item.Substring(position);
                double dArea;
                double.TryParse(area,out dArea);
                switch (ground)
                { 
                    case "水田":
                        land.Paddy = dArea;
                        break;
                    case "水浇地":
                        land.Irrigated = dArea;
                        break;
                    case "旱地":
                        land.Dry = dArea;
                        break;
                    default: break;
                }
            }
            return land;
        }
    }
}