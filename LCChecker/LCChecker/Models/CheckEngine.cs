using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace LCChecker.Models
{
    public class CheckEngine
    {
        /// <summary>
        /// 验证规则
        /// </summary>
        public List<RuleInfo> rules = new List<RuleInfo>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();

       // public Dictionary<string, string> Error2 = new Dictionary<string, string>();

        /// <summary>
        /// 提示
        /// </summary>
        public Dictionary<string, string> Warning = new Dictionary<string, string>();

        /// <summary>
        /// 存放自查表中的一些数据
        /// </summary>
        public Dictionary<string, Index2> Ship = new Dictionary<string, Index2>();

        /// <summary>
        /// 检查表格中的项目编号们
        /// </summary>
        public List<string> IDS = new List<string>();


        /// <summary>
        /// 用于验证 表4  表5  表8 跟表3之间关系  编号 bool
        /// </summary>
        public Dictionary<string, bool> Whether = new Dictionary<string, bool>();

        public Dictionary<string, List<string>> GetError()
        {
            return Error;
        }

        public Dictionary<string, string> GetWarning()
        {
            return Warning;
        }

        public List<string> GetIDS()
        {
            return IDS;
        }
        public virtual void SetWhether(List<Project> projects)
        {

        }

        /// <summary>
        /// 验证项目编号格式
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool VerificationID(string value)
        {
            return Regex.IsMatch(value, @"^33[0-9]{12}");
        }
        public bool Check(string filePath, ref string mistakes, ReportType Type, List<Project> Data, bool flag)
        {
            if (flag)
            {
                return Check(filePath, ref mistakes, Type, Data);
            }
            else
            {
                return Check(filePath, ref mistakes, Type);
            }
        }

        /// <summary>
        ///  用于 4 5 8 报部表格检查
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="mistakes"></param>
        /// <param name="Type"></param>
        /// <param name="Data"></param>
        /// <returns></returns>
        public bool Check(string filePath, ref string mistakes, ReportType Type, List<Project> Data)
        {
            int startRow = 0, startCell = 0;
            ISheet sheet = OpenSheet(filePath, true, ref startRow, ref startCell, ref mistakes, Type);
            if (sheet == null)
            {
                if (Error.ContainsKey("表格格式内容"))
                {
                    Error["表格格式内容"].Add("提交的表格无法检索 请核对格式");
                }
                else {
                    Error.Add("表格格式内容", new List<string> { "提交的表格无法检索 请核对格式" });
                }
                
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
                if (!VerificationID(value))
                    continue;
                if (IDS.Contains(value))
                {
                    if (Error.ContainsKey(value))
                    {
                        Error[value].Add("表格中存在相同项目编号");
                    }
                    else
                    {
                        Error.Add(value, new List<string>() { "表格中存在相同项目编号" });
                    }
                    continue;
                }
                IDS.Add(value);
                if (Whether.ContainsKey(value))//在表3中存在
                {
                    if (Whether[value])//重点复核确认总表中 填：是  提交表格存在这个项目  检查这个项目填写数据与格式
                    {
                        foreach (var item in rules)
                        {
                            if (!item.Rule.Check(row, startCell))
                            {
                                ErrorRow.Add(item.Rule.Name);
                            }
                        }
                        if (ErrorRow.Count() != 0)
                        {
                            if (Error.ContainsKey(value))
                            {
                                Error[value] = ErrorRow;
                            }
                            else {
                                Error.Add(value, ErrorRow);
                            }
                            
                        }
                    }
                    else
                    {//重点复核确认总表中 填：否  提交表格中存在   处理：提示
                        Warning[value] = "规则0006： 与重点项目复核确认总表中项目类型不符";
                        //Warning.Add(value, "与重点项目复核确认总表中项目类型不符");
                    }
                    Whether.Remove(value);
                }
                else
                {//重点复核确认总表中 没有这个项目  提交表格中存在  处理：错误
                    ErrorRow.Add("规则0007：重点复核确认总表中不存在项目");
                    if (!Error.ContainsKey(value))
                    {
                       Error.Add(value, ErrorRow);
                        //Error2.Add(value, "重点复核确认总表中不存在项目");
                        //Error2[value] = "重点复核确认总表中不存在项目";
                    }
                        
                }

            }
            foreach (var item in Whether.Keys)
            {
                if (Whether[item])//重点复核确认总表中 填：是  提交表格中没有这个项目  处理：提示
                {
                    Warning[item] = "规则0005：项目存在于重点项目复核确认总表，但不存在与本表中。";
                    //Warning.Add(item, "项目存在于重点项目复核确认总表，但不存在于本表中。");
                }
            }

            return true;
        }


        





        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="mistakes"></param>
        /// <param name="Type"></param>
        /// <returns></returns>
        public bool Check(string filePath, ref string mistakes, ReportType Type)
        {
            int startRow = 0, startCell = 0;
            ISheet sheet = OpenSheet(filePath, true, ref startRow, ref startCell, ref mistakes, Type);
            if (sheet == null)
            {
                return false;
            }
            startRow++;
            for (int i = startRow; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                var value = row.GetCell(3+startCell, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                if (!VerificationID(value))
                    continue;
                List<string> ErrorRow = new List<string>();
                if (IDS.Contains(value))
                {
                    if (Error.ContainsKey(value))
                    {
                        Error[value].Add("表格中存在相同的项目");
                    }
                    else {
                        Error.Add(value, new List<string>{"表格中存在相同的项目"});
                    }
                    continue;
                }
                IDS.Add(value);
                foreach (var item in rules)
                {
                    if (!item.Rule.Check(row, startCell))
                    {
                        ErrorRow.Add(item.Rule.Name);
                    }
                }
                if (ErrorRow.Count != 0)
                {
                    Error.Add(value, ErrorRow);
                }
            }
            return true;
        }



        /// <summary>
        /// 检查表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="startCell"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public bool FindHeader(ISheet sheet, ref int startRow, ref int startCell, ReportType type)
        {

            var Name = @"([\w\W])" + type.GetDescription();
            string[] Header = { "编号", "市", "县" };
            for (int i = 0; i < 20; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < 20; j++)
                    {
                        var value = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                        if (string.IsNullOrEmpty(value))
                            continue;
                        if (value == type.ToString())
                        {
                            var Row = sheet.GetRow(i + 1);
                            if (Row == null)
                                return false;
                            value = Row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                            if (string.IsNullOrEmpty(value))
                                return false;
                            if (!Regex.IsMatch(value, Name))
                                return false;
                            i = i + 4;
                            row = sheet.GetRow(i);
                            if (row == null)
                                return false;
                            value = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                            if (string.IsNullOrEmpty(value))
                                return false;
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
                            return false;
                        }
                    }
                }
            }
            return false;

        }




        /// <summary>
        /// 打开Excel，找第一个Sheet，并返回表头
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="findHeader"></param>
        /// <returns></returns>
        public ISheet OpenSheet(string filePath, bool findHeader, ref int startRow, ref int startCol, ref string errMsg, ReportType Type)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch(Exception er)
            {
                string str = er.ToString();
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

            if (FindHeader(sheet, ref startRow, ref startCol, Type) == false)
            {
                errMsg = "未找到附表文件的表头";
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
            int startRow = 1, starCell = 0;
            ISheet sheet = OpenSheet(filePath, false, ref startRow, ref starCell, ref fault, ReportType.附表8);
            if (sheet == null)
            {
               // Error2["表格内容格式"] = "获取项目总表失败";
                //Error2.Add("表格内容格式", "获取项目总表失败");
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
                var CurrentData = GetCurrentData(row);
                if (CurrentData == null)
                {
                    Error.Add(value, new List<string> { "获取核对数据失败，导致无法验证" });
                }
                else {
                    Ship[value] = CurrentData;     
                }
                
            }
        }



        /// <summary>
        /// 获取补充耕地项目中的数据
        /// </summary>
        /// <param name="filePath"></param>
        public void GetSuppleMessage(string filePath)
        {
            string fault = "";
            int startRow = 1, startCell = 0;
            ISheet sheet = OpenSheet(filePath, false, ref startRow, ref startCell, ref fault, ReportType.附表8);
            if (sheet == null)
            {
                //Error2["项目总表"] = "获取项目总表失败";
             //   Error2.Add("项目总表", "获取项目总表失败");
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
                var CurrentData = GetCurrentData(row);
                if (CurrentData == null)
                {
                    //Error2["初始化"] = "获取核对信息失败";
                    //Error2.Add("初始化", "获取核对信息失败");
                    Error.Add("初始化", new List<string> { "获取核对信息失败" });
                }
                else {
                    Ship[value] = CurrentData;
                    //Ship.Add(value, CurrentData);
                }
                

            }
        }


        public Index2 GetCurrentData(NPOI.SS.UserModel.IRow row)
        {
            bool Flag = false;
            var YesOrNo = row.GetCell(9, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            if (string.IsNullOrEmpty(YesOrNo) || YesOrNo == "否")
            {
                Flag = false;
            }
            else
            {
                Flag = true;
            }
            //等别
            var CellGrade = row.GetCell(35, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            double grade;
            if (CellGrade.CellType == CellType.Numeric || CellGrade.CellType == CellType.Formula)
            {
                try
                {
                    grade = CellGrade.NumericCellValue;
                }
                catch
                {
                    grade = .0;
                }
            }
            else {
                var valGrade = CellGrade.ToString().Trim();
                double.TryParse(valGrade, out grade);
            }

            //表8  14栏
            var cell = row.GetCell(13, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            double data2;
            if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
            {
                try
                {
                    data2 = cell.NumericCellValue;
                }
                catch
                {
                    data2 = .0;
                }

            }
            else
            {
                var Indicator = cell.ToString();
                double.TryParse(Indicator, out data2);
            }
            var cell2 = row.GetCell(19, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            double twenty;
            if (cell2.CellType == CellType.Numeric || cell2.CellType == CellType.Formula)
            {
                try
                {
                    twenty = cell2.NumericCellValue;
                }
                catch
                {
                    twenty = .0;
                }
            }
            else
            {
                var val2 = cell2.ToString();
                double.TryParse(val2, out twenty);
            }
            var cell3 = row.GetCell(27, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            double TwentyEight;
            if (cell3.CellType == CellType.Numeric || cell3.CellType == CellType.Formula)
            {
                try
                {
                    TwentyEight = cell3.NumericCellValue;
                }
                catch
                {
                    TwentyEight = .0;
                }
            }
            else
            {
                var val3 = cell3.ToString().Trim();
                double.TryParse(val3, out TwentyEight);
            }
            var val = row.GetCell(36, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();

            Land LandData = GetLand(val);

            Index2 CurrentData = new Index2()
            {
                IsApplyDelete = Flag,
                Grade = grade,//质量等别
                Indicators = data2 + twenty + TwentyEight,//  表 8  14栏
                Land = LandData//  水田  水浇地  旱地  数据
            };
            return CurrentData;
        }

        public Land GetLand(string value)
        {
            Land land = new Land();
            string val = value.Replace("，", ",").Replace("。", string.Empty);
            var team = val.Split(',');
            foreach (string item in team)
            {
                Regex r = new Regex(@"-?[0-9]");
                string data = r.Match(item).ToString();
                int position = item.IndexOf(data);
                if (position == 0)
                    continue;
                string ground = item.Substring(0, position);
                string area = item.Substring(position);
                double dArea;
                double.TryParse(area, out dArea);
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