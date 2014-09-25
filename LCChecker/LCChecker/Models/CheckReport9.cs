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
      
        public CheckReport9(string filePath,List<Project> projects)
        {
            GetMessage(filePath);
            Dictionary<string, Project> Team = new Dictionary<string, Project>();
            foreach (var item in projects)
            {
                Team.Add(item.ID, item);
            }
            var list = new List<IRowRule>();
            list.Add(new OnlyProject() { ColumnIndex = 3, Projects = Team, Values = new[] { "项目编号", "市", "县", "项目名称", "新增耕地面积" } });
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
                    Mistakes = "未找到水田  水浇地  旱地列";
                    return false;
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
                Land MineLand = new Land() { Paddy=Area[0],Irrigated=Area[1],Dry=Area[2]};
                if (Ship.ContainsKey(value))
                {
                    Index2 Data = Ship[value];
                    if (!Data.Land.Compare(MineLand))
                    {
                        ErrorRow.Add("水田、水浇地、旱地与自检表中数据不符");
                    }
                    if (Degree1[1] != 0)
                    {
                        ErrorRow.Add("水浇地存在问题");
                    }
                    double CurrentDegree;
                    double.TryParse(Data.Grade, out CurrentDegree);
                    if (((Degree1[0] - CurrentDegree) * (CurrentDegree - Degree1[2])) < 0)
                    {
                        ErrorRow.Add("耕地质量等别与自检表不符");
                    }
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
                    Mistakes = "存在填写多个质量等别";
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