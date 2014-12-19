using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public class CheckNine:SecondCheckEngine,ISeCheck
    {

        
        public Dictionary<string, SecondProject> Team;
        public CheckNine(List<SecondProject> projects) {
            Team= projects.ToDictionary(e => e.ID, e => e);
            var list = new List<IRowRule>();
            //list.Add(new OnlySecondProject() { ColumnIndex = 3, NewAreaIndex = 2, Projects = Team, ID = "2901（基本规则）", Values = new[] { "市", "县", "项目名称", "新增耕地面积" } });
            list.Add(new OnlySecondProject() { ColumnIndex = 3, NewAreaIndex = 2, Projects = Team, ID = "2901（基本规则）", Values = new[] { "市", "县", "项目名称" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 22, Values = new[] { "是", "否" } ,ID="2902（填写规则）"});

            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

        public new bool Check(string FilePath, ref string Mistakes, SecondReportType Type, bool IsPlan)
        {
            return CheckSpecial(FilePath, ref Mistakes, Type);
        }


        public bool CheckSpecial(string FilePath, ref string Mistakes, SecondReportType Type) {
            int StartRow = 0, StartCell = 0;
            ISheet sheet = XslHelper.OpenSheet(FilePath, true, ref StartRow, ref StartCell, ref Mistakes, Type);
            if (sheet == null)
            {
                Mistakes = "检索表格内无数据";
                return false;
            }
            StartRow++;
            IRow checkRow = sheet.GetRow(StartRow);
            if (!CheckDegree(checkRow, StartCell)) {
                Mistakes = "附表9项目新增耕地具体情况表格格式请核对";
                return false;
            }
            int Max = sheet.LastRowNum;
            for (var i = StartRow+1; i <= Max; i=i+3) {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    break;
                var value = row.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                if (!value.VerificationID())
                    continue;
                if (!XslHelper.JudgeLand(sheet, i, StartCell)) {
                    Mistakes = value + "规则000：未找到水田、水浇地、旱地列";
                    continue;
                }
                if (IDS.Contains(value))
                {
                    if (Error.ContainsKey(value))
                    {
                        Error[value].Add("规则0001：表格中存在相同项目");
                    }
                    else
                    {
                        Error.Add(value, new List<string> { "规则0001：表格中存在相同项目" });
                    }
                    continue;
                }
                IDS.Add(value);
                string Fault = "";
                int[] Degree1 = new int[3];
                double[] Area = new double[3];
                List<string> ErrorRow = new List<string>();
                for (var j = 0; j < 3; j++) {
                    Fault = "";
                    if (!XslHelper.CheckLand(sheet, i + j, ref Area[j], ref Degree1[j], ref Fault, StartCell + 7)) {
                        ErrorRow.Add(Fault);
                    }
                }
                if (Team.ContainsKey(value))
                {
                    var sum = Area[0] + Area[2];
                    var areacell = row.GetCell(StartCell + 5, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    double Barea = 0.0;
                    if (areacell.CellType == CellType.Numeric || areacell.CellType == CellType.Formula)
                    {
                        try
                        {
                            Barea = areacell.NumericCellValue;
                        }
                        catch
                        {
                            Barea = 0.0;
                        }
                    }
                    else
                    {
                        double.TryParse(areacell.ToString().Trim(), out Barea);
                    }
                    if (Math.Abs(Barea - sum) > 0.0001) {
                        ErrorRow.Add("规则2902（数据规则）：水田、旱地的面积之和与复核确认验收项目清单新增耕地面积不符");
                    }
                    if (Area[1] != 0 && Degree1[1] != 0)
                    {
                        ErrorRow.Add("错误0091：水浇地栏错误");
                    }

                    foreach (var item in rules)
                    {
                        if (!item.Rule.Check(row, StartCell))
                        {
                            ErrorRow.Add(item.Rule.Name);
                        }
                    }
                }
                else {
                    ErrorRow.Add("规则0002（一致性）：复核确认验收项目清单中不存在该项目，请核对");
                }
                

                //if (Team.ContainsKey(value))
                //{
                //    var currentProject = Team[value];
                //    var sum = Area[0] + Area[2];
                //    if (currentProject.NewArea.HasValue)
                //    {
                //        if (Math.Abs(currentProject.NewArea.Value - sum) > 0.0001)
                //        {
                //            ErrorRow.Add("规则2902（数据规则）：水田、旱地的面积之和与复核确认验收项目清单新增耕地面积不符");
                //        }
                //    }
                //    else
                //    {
                //        ErrorRow.Add("错误0000：数据库中没有新增耕地面积值，无法进行水田、旱地面积核对");
                //    }
                //    if (Area[1] != 0 && Degree1[1] != 0)
                //    {
                //        ErrorRow.Add("错误0091：水浇地栏错误");
                //    }

                //    foreach (var item in rules) {
                //        if (!item.Rule.Check(row, StartCell)) {
                //            ErrorRow.Add(item.Rule.Name);
                //        }
                //    }
                //}
                //else {
                //    ErrorRow.Add("规则0002（一致性）：复核确认验收项目清单中不存在该项目，请核对");
                //}

                if (ErrorRow.Count() != 0)
                {
                    if (Error.ContainsKey(value))
                    {
                        Error[value].Add("表格中存在相同项目");
                    }
                    else
                    {
                        Error.Add(value, ErrorRow);
                    }
                }
                else
                {
                    if (Degree1[0] != 0 && (Math.Abs(Area[0]-0) > 0.0001))
                    {
                        if (!DicPaddy.ContainsKey(value))
                        {
                            DicPaddy.Add(value, new SeLand() { Degree = (Degree)Degree1[0], Area = Area[0] });
                        }
                    }
                    if (Degree1[2] != 0 && (Math.Abs(Area[2]-0) > 0.0001)) {
                        if(!DicDry.ContainsKey(value)){
                            DicDry.Add(value, new SeLand() { Degree = (Degree)Degree1[2], Area = Area[2] });
                        }
                    }
                }

            }


            return true;

        }

        public bool CheckDegree(IRow row, int StartCell) {
            string[] Degrees = {"一等","二等","三等","四等","五等","六等","七等","八等","九等","十等","十一等","十二等","十三等","十四等","十五等" };
            int Serial=StartCell+7;
            foreach (var item in Degrees) {
                var value = row.GetCell(Serial++).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    return false;
                if (value.ToLower() != item.ToLower())
                    return false;
            }
            return true;
        }
    }
}