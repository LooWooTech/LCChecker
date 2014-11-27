using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public class CheckEight:SecondCheckEngine,ISeCheck
    {
       // public Dictionary<HookedProject, List<Dictionary<BuildProject, List<ReHookProject>>>> Data = new Dictionary<HookedProject, List<Dictionary<BuildProject, List<ReHookProject>>>>();

        IRowRule ReHookRule = new SumRowRule() { SumColumnIndex = 21, ColumnIndices = new[] { 22, 23 } };
        public CheckEight(List<SecondProject> projects) {

            Whether = projects.ToDictionary(e => e.ID, e => e.IsRelieve);
            
            Dictionary<string, SecondProject> Team = projects.ToDictionary(e => e.ID, e => e);
            
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, Projects = Team, NewAreaIndex = 2, Values = new[] { "市", "县", "项目名称", "新增耕地面积" }, ID = "2801" });
            list.Add(new NoLessThanRowRule() { Column1Index = 7, Column2Index = 8, ID = "2802" });

            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

        public new bool Check(string FilePath, ref string Mistakes, SecondReportType Type) {
            return CheckSpecial(FilePath, ref Mistakes, Type);
        }

        public bool CheckSpecial(string FilePath, ref string Mistakes, SecondReportType Type) {
            int StartRow = 0, StartCell = 0;
            ISheet sheet = XslHelper.OpenSheet(FilePath, true, ref StartRow, ref StartCell, ref Mistakes, Type);
            if (sheet == null)
                return false;
            StartRow++;
            int Max = sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++) {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    continue;
                var Key1 = row.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(Key1))
                    continue;
                if (!Key1.VerificationID())
                    continue;
                if (IDS.Contains(Key1)) {
                    if (Error.ContainsKey(Key1))
                    {
                        Error[Key1].Add("规则000：表格中存在相同的项目");
                    }
                    else {
                        Error.Add(Key1, new List<string> { "规则000：表格中存在相同的项目" });
                    }
                    continue;
                }
                IDS.Add(Key1);

                List<string> ErrorRow = new List<string>();
                if (Whether.ContainsKey(Key1))
                {
                    if (!Whether[Key1])
                    {
                        ErrorRow.Add("规则：与重点复核确认项目以外所有报部备案项目复核确认总表不符");
                        //Warning[Key1] = "提示：与重点复核确认项目以外所有报部备案项目复核确认总表不符";
                    }
                    
                    foreach (var item in rules) {
                        if (!item.Rule.Check(row, StartCell)) {
                            ErrorRow.Add(item.Rule.Name);
                        }
                    }
                    Whether.Remove(Key1);
                }
                else {
                    ErrorRow.Add("规则000：与复核确认验收项目清单不符，该项目不存在");
                }
                double Nine = 0.0;
                double.TryParse(row.Cells[StartCell + 8].GetValue().ToString(), out Nine);
                int spanR1 = 0,spanC1=0;
                if (XslHelper.isMergeCell(sheet, i+1, StartCell+1, out spanR1, out spanC1)) {
                    int FlagR = spanR1;
                    int Rowoffset = 0;
                    double Sumfifteent = 0.0;
                    while(FlagR>0){
                        int spanR2 = 0, spanC2 = 0;
                        if (XslHelper.isMergeCell(sheet, i + 1 + Rowoffset, StartCell + 10, out spanR2, out spanC2))
                        {

                            var row2 = sheet.GetRow(i + Rowoffset);
                            if (row2 == null)
                            {
                                ErrorRow.Add("未获取对应建设用地项目情况");
                                break;
                            }
                            double fifteen = 0.0;
                            double.TryParse(row2.Cells[StartCell + 14].GetValue().ToString(), out fifteen);
                            Sumfifteent += fifteen;

                            double sumtwenty = 0.0;
                            for (var j = 0; j < spanR2; j++)
                            {
                                IRow row3 = sheet.GetRow(j + i + Rowoffset);
                                if (row3 == null)
                                {
                                    ErrorRow.Add("未获取重新不挂补充耕地项目情况");
                                    break;
                                }
                                if (!ReHookRule.Check(row3, StartCell))
                                {
                                    ErrorRow.Add(ReHookRule.Name);
                                }
                                double Twenty = 0.0;
                                if (double.TryParse(row3.Cells[StartCell + 22].GetValue().ToString(), out Twenty))
                                {
                                    sumtwenty += Twenty;
                                }
                                else {
                                    var key3 = row3.Cells[StartCell + 18].GetValue().ToString();
                                    ErrorRow.Add(key3 + ":无法获取重新补挂补充耕地项目的挂钩占补平衡面积");
                                }
                            }

                            if (Math.Abs(sumtwenty - fifteen) > 0.0001) {
                                ErrorRow.Add("规则000：对应建设用地项目的解挂可用于占补平衡面积等于重新补挂补充耕地项目的用于建设项目挂钩可用于占补平衡面积和");
                            }

                            Rowoffset += spanR2;
                            FlagR -= spanR2;
                        }
                        else {
                            ErrorRow.Add("无法获取合并单元格信息");
                            break;
                        }
                    }
                    if (Math.Abs(Sumfifteent - Nine) > 0.0001) {
                        ErrorRow.Add("规则000：已挂钩补充耕地项目的需解除已挂钩使用可用于占补平衡面积等于对应建设用地项目的用于该建设项目解挂可用于占补平衡面积和");
                    }
                    
                }
                i = i + spanR1 - 1;
                if (ErrorRow.Count() != 0) {
                    Error[Key1] = ErrorRow;
                }
                


            }

            foreach (var item in Whether.Keys) {
                if (Whether[item]) {
                    if (Error.ContainsKey(item))
                    {
                        Error[item].Add("规则000：与重点复核确认项目以外所有报部备案项目复核确认总表不符,该项目应存在与本表中");
                    }
                    else {
                        Error.Add(item, new List<string> { "规则000：与重点复核确认项目以外所有报部备案项目复核确认总表不符,该项目应存在与本表中" });
                    }
                    
                   // Warning[item] = "提示：与重点复核确认项目以外所有报部备案项目复核确认总表不符,该项目应存在与本表中";
                }
            }
                
            return true;
        }



        














        //public int GetLength(ISheet sheet,int RowIndex,int Index) {
        //    int length = 0;
        //    int Max=sheet.LastRowNum;
        //    for (var i = RowIndex; i <= Max; i++) {
        //        IRow row = sheet.GetRow(i);
        //        if (row == null)
        //            break;
        //        length++;
        //        var value = row.Cells[Index].GetValue().ToString().Trim();
        //        if (string.IsNullOrEmpty(value))
        //            continue;
        //        if (value.VerificationID())
        //            break;
        //    }
        //    return length - 1;
        //}

        //public bool Get(string FilePath, ref string Mistakes, SecondReportType Type)
        //{
        //    List<ReHookProject> rehookproject=new List<ReHookProject>();
        //    Dictionary<BuildProject, List<ReHookProject>> Dicbuildproject=null;
        //    List<Dictionary<BuildProject, List<ReHookProject>>> buildprojects=null;
        //    int StartRow = 0, StartCell = 0;
        //    ISheet sheet = XslHelper.OpenSheet(FilePath, true, ref StartRow, ref StartCell, ref Mistakes, Type);
        //    if (sheet == null)
        //        return false;
        //    StartRow++;
        //    int Max = sheet.LastRowNum;
        //    BuildProject Key2=null;
        //    HookedProject Key1;
        //    for (var i = StartRow; i <= Max; i++) {
        //        IRow row = sheet.GetRow(i);
        //        if (row == null)
        //            continue;
        //        var value3 = row.GetCell(StartCell + 18, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
        //        if (string.IsNullOrEmpty(value3))
        //            continue;
        //        if (!value3.VerificationID())
        //            continue;
        //        var value2 = row.GetCell(StartCell + 12, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
        //        var value1 = row.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
        //        City city = 0;
        //        double[] Area = new double[4];

        //        if (!string.IsNullOrEmpty(value1) && value1.VerificationID()) {
        //            for (var j = 0; j < 4; j++) {
        //                double.TryParse(row.Cells[StartCell + 5 + j].GetValue().ToString(), out Area[j]);
        //            }
        //            if (Enum.TryParse<City>(row.Cells[StartCell + 1].GetValue().ToString(), out city)) {
        //                HookedProject hookedproject = new HookedProject()
        //                {
        //                    City = city,
        //                    County = row.Cells[StartCell + 2].GetValue().ToString().Trim(),
        //                    ID = value1,
        //                    Name = row.Cells[StartCell + 4].GetValue().ToString().Trim(),
        //                    NewArea = Area[0],
        //                    Balance = Area[1],
        //                    HookedBalance = Area[2],
        //                    RelieveBalance = Area[3]
        //                };
        //                Key1 = hookedproject;
        //            } 
        //            Dicbuildproject.Clear();
        //            rehookproject.Clear();
        //        }


        //        if (!string.IsNullOrEmpty(value2) && value2.VerificationID()) {
        //            if (Key2 != null)
        //            {
        //                Dicbuildproject = new Dictionary<BuildProject, List<ReHookProject>>();
        //                Dicbuildproject[Key2] = rehookproject;
        //                buildprojects.Add(Dicbuildproject);
        //            }
                    






        //            double.TryParse(row.Cells[StartCell + 14].GetValue().ToString(), out Area[0]);
        //            if(Enum.TryParse<City>(row.Cells[StartCell+10].GetValue().ToString(),out city)){
        //                BuildProject buildproject = new BuildProject()
        //                {
        //                    City=city,
        //                    County=row.Cells[StartCell+11].GetValue().ToString().Trim(),
        //                    ID=value2,
        //                    Name=row.Cells[StartCell+13].GetValue().ToString().Trim(),
        //                    BuildBalance=Area[0]
        //                };
        //                Key2 = buildproject;
        //            }
        //            rehookproject.Clear();
                    
        //        }
                
        //        for (var j = 0; j < 4; j++) {
        //            double.TryParse(row.Cells[StartCell + 20 + j].GetValue().ToString(), out Area[j]);
        //        }
                
        //        if (Enum.TryParse<City>(row.Cells[StartCell + 16].GetValue().ToString(), out city)) {
        //            rehookproject.Add(new ReHookProject
        //            {
        //                City = city,
        //                County = row.Cells[StartCell + 17].GetValue().ToString(),
        //                ID = value3,
        //                Name = row.Cells[StartCell + 19].GetValue().ToString().Trim(),
        //                AvailBalance = Area[0],
        //                BeBalance = Area[1],
        //                BuildBalance = Area[2],
        //                AfBalance = Area[3]
        //            });
        //        }








 

        //    }
        //        return true;
        //}

        //public void GetBuildProjects(ISheet sheet, int StartRow, int offset) { 
            
        //}

        //public List<ReHookProject> GetReHookProjects(ISheet sheet,int StartRow,int offset,int Length) {
        //    List<ReHookProject> projects=new List<ReHookProject>();
        //    int Max=StartRow+Length;
        //    for (var i = StartRow; i < Max; i++)
        //    {
        //        IRow row = sheet.GetRow(i);
        //        if (row == null)
        //            continue;
        //        var value = row.GetCell(offset + 18, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
        //        if (string.IsNullOrEmpty(value))
        //            continue;
        //        if (!value.VerificationID())
        //            continue;
        //        City city = 0;
        //        if (Enum.TryParse<City>(row.Cells[offset + 16].GetValue().ToString(), out city)) {
        //            double availbalance = 0, bebalance = 0, buildbalance = 0, afbalance = 0;
        //            double.TryParse(row.Cells[offset + 20].GetValue().ToString(), out availbalance);
        //            double.TryParse(row.Cells[offset + 21].GetValue().ToString(), out bebalance);
        //            double.TryParse(row.Cells[offset + 22].GetValue().ToString(), out buildbalance);
        //            double.TryParse(row.Cells[offset + 23].GetValue().ToString(), out afbalance);
        //            projects.Add(new ReHookProject
        //            {
        //                City = city,
        //                County = row.Cells[offset + 17].GetValue().ToString(),
        //                ID = value,
        //                Name = row.Cells[offset + 19].GetValue().ToString().Trim(),
        //                AvailBalance = availbalance,
        //                BeBalance = bebalance,
        //                BuildBalance = buildbalance,
        //                AfBalance = afbalance
        //            });
        //        }
        //    }
                
        //    return projects;
        //}


        

    }
}