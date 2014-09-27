using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace LCChecker.Models
{


    public class CheckReport9:CheckEngine ,ICheck
    {
        public Dictionary<string, Project> Team = new Dictionary<string, Project>();

        public CheckReport9(string filePath,List<Project> projects)
        {
            GetMessage(filePath);
            foreach (var item in projects)
            {
                Team.Add(item.ID, item);
            }
            var list = new List<IRowRule>();
            list.Add(new OnlyProject() { ColumnIndex = 3, Projects = Team, Values = new[] { "项目编号", "市", "县", "项目名称", "新增耕地面积" },ID="2901" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 22, Values = new[] { "是", "否" } ,ID="2903"});

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
            for (int i = startRow+1; i <= sheet.LastRowNum; i=i+3)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                var value = row.GetCell(startCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                if (!VerificationID(value))
                    continue;
                if (!JudgeLand(sheet, i, startCell))
                {
                    Mistakes = "规则2907：未找到水田  水浇地  旱地列";
                    return false;
                }
                if (IDS.Contains(value))
                {
                    if (Error.ContainsKey(value))
                    {
                        Error[value].Add("存在相同项目");
                        continue;
                    }
                    else {
                        Error.Add(value, new List<string>(){"存在相同项目"});
                    }
                }
                else {
                    IDS.Add(value);
                }
                string Fault="";

                int[] Degree1 = new int[3];
                double[] Area = new double[3];

                List<string> ErrorRow = new List<string>();
                for (var j = 0; j < 3; j++)
                {
                    Fault = "";
                    if (!CheckLand(sheet, i+j, ref Area[j], ref Degree1[j], ref Fault, startCell + 7))
                    {
                        ErrorRow.Add(Fault);
                    }
                }
                if (Ship.ContainsKey(value))//ship字典中包含了自查表中所有项目  与项目耕地质量等别、第10栏 填‘是’‘否’的关系
                {
                    if (Team.ContainsKey(value))
                    {
                        var CurrentProject = Team[value];
                        var sum = Area[0] + Area[2];
                        if (CurrentProject.NewArea.HasValue)
                        {
                            if (Math.Abs(CurrentProject.NewArea.Value - sum) > 0.0001)
                            {
                                ErrorRow.Add("水田、旱地的面积之和与自查表新增耕地面积不符");
                            }
                        }
                        else {
                            ErrorRow.Add("数据库中没有新增耕地面积的值，无法进行水田、旱地面积核对");
                        }
                        
                    }

                    Index2 Data = Ship[value];
                    if (!Data.IsApplyDelete)//该项目在自查表中填写 ‘否’或者不填  同时也在附表9中出现，那么就remove该项目编号key
                    {
                        Ship.Remove(value);
                    }
                    if (Degree1[1] != 0)
                    {
                        ErrorRow.Add("水浇地存在问题");
                    }
                    double CurrentDegree;
                    double.TryParse(Data.Grade, out CurrentDegree);
                    if (((Degree1[0] - CurrentDegree) * (CurrentDegree - Degree1[2])) < 0)
                    {
                        ErrorRow.Add("规则2906：自查表耕地质量等别在水田旱地等别之间，可以与其中一个等别相等");
                    }
                }
                else {
                    Error.Add(value, new List<string>() {"自查表中不存在该项目，请核对" });
                    continue;
                }
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
            foreach (var item in Ship.Keys)
            {
                if (!Ship[item].IsApplyDelete)
                {
                    Warning.Add(item, "规则0005：项目与自查表不符，该项目应该包括在本表中。");
                }
            }
            
            return true;
        }


        public bool JudgeLand(NPOI.SS.UserModel.ISheet sheet,int Line,int xoffset=0)
        {
            string[] Lands = new string[] { "水田", "水浇地", "旱地" };
            foreach (var item in Lands)
            {
                IRow row = sheet.GetRow(Line++);
                if (row == null)
                {
                    return false;
                }
                var value = row.GetCell(6 + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    return false;
                if (value != item)
                    return false;
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


        public bool CheckLand(NPOI.SS.UserModel.ISheet sheet,int Line,ref double LandArea,ref int Degree,ref string Mistakes,int xoffset=7)
        {
            IRow row = sheet.GetRow(Line);
            if (row == null)
            {
                Mistakes = "未获得相关表格行";
                return false;
            }
            int Max=xoffset+15;
            bool Flag = false;
            double Area = 0.0;
            for (var i = xoffset; i < Max; i++)
            {
                var cell = row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (string.IsNullOrEmpty(cell.ToString().Trim()))
                    continue;
                if (Flag)
                {
                    Mistakes = "规则2904：水田、水浇地、旱地中同一个分类不允许填写多个质量等别";
                    return false;
                }
                
                if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                {
                    try
                    {
                        Area = cell.NumericCellValue;
                    }
                    catch
                    {
                        Area = .0;
                    }
                }
                else {
                    var val = cell.ToString().Trim();
                    double.TryParse(val, out Area);
                }
                Degree = i-6;
                Flag = true;
            }
            Area = Area / 15;
            LandArea = Math.Floor(Area * 10000) / 10000;   
            return true;
        }


        public new  bool Check(string FilePath, ref string Mistakes,ReportType Type,List<Project> Data,bool flag)
        {
            return CheckSpecial(FilePath,ref Mistakes,Type);
        }

    }


   
}