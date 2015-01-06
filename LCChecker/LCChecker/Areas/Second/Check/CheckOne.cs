﻿using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Areas.Second
{
    public class CheckOne:SecondCheckEngine,ISeCheck,ISePlanCheck
    {        
        public CheckOne(List<SecondProject> projects) {
            
            Whether = projects.ToDictionary(e => e.ID, e => true);
            
            Dictionary<string,SecondProject> Team=projects.ToDictionary(e=>e.ID,e=>e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, AreaIndex = 0, NewAreaIndex = 0, Projects = Team, ID = "2101(基本规则)", Values = new[] { "项目名称", "市", "县" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是", "否" }, ID = "2102（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是", "否" }, ID = "2103（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" }, ID = "2104（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2105（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" }, ID = "2106（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 10, Values = new[] { "是", "否" }, ID = "2107（填写规则）" });
            list.Add(new ConditionalRowRule()
            {
                ID = "2108（填写规则）",
                Condition = new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是" } },
                Rule = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "否" } }
            });
            list.Add(new ConditionalRowRule()
            {
                ID="2109（填写规则）",
                Condition = new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是" } },
                Rule = new Less() { ColumnIndex = new[] { 6, 7, 8, 9, 10 }, Value = "是" }
            });
            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }

        public CheckOne(List<pProject> projects) {

            //Whether = projects.ToDictionary(e => (e.Name.Trim().ToUpper() + '-' + e.County.Trim().ToUpper() + '-' + e.Key.Trim().ToUpper()), e => true);
            foreach (var item in projects) {
                var key = item.Name.Trim().ToUpper() + '-' + item.County.Trim().ToUpper() + '-' + item.Key.Trim().ToUpper();
                if (PlanIDS.ContainsKey(key))
                {
                    PlanIDS[key]++;
                }
                else {
                    PlanIDS.Add(key, 1);
                }
            }
            var list = new List<IRowRule>();
            list.Add(new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是", "否" }, ID = "2102（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是", "否" }, ID = "2103（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" }, ID = "2104（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2105（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" }, ID = "2106（填写规则）" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 10, Values = new[] { "是", "否" }, ID = "2107（填写规则）" });
            list.Add(new ConditionalRowRule()
            {
                ID = "2108（填写规则）",
                Condition = new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是" } },
                Rule = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "否" } }
            });
            list.Add(new ConditionalRowRule()
            {
                ID = "2109（填写规则）",
                Condition = new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是" } },
                Rule = new Less() { ColumnIndex = new[] { 6, 7, 8, 9, 10 }, Value = "是" }
            });
            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }


        public new bool Check(string FilePath, ref string Mistakes, SecondReportType Type,bool IsPlan) {
            //检查上传的报部表格   附表一主要检查 市、县、名称、  市级自查是否存在疑问   属于申请删除  等几栏填写  是  否
            if (IsPlan)
            {
                if (!PlanCheckEngine(FilePath, ref Mistakes, Type)) {
                    return false;
                }
                if (!GetPlanProject(FilePath, ref Mistakes, Type)) {
                    return false;
                }
            }
            else {
                if (!CheckEngine(FilePath, ref Mistakes, Type))
                {
                    return false;
                }
                //获取用户上传的附表1中的正确数据
                if (!GetProject(FilePath, ref Mistakes, Type))
                {
                    return false;
                }
            }
            return true;
        }

        public bool GetPlanProject(string FilePath, ref string Mistakes, SecondReportType Type) {
            int StartRow = 0, StartCell = 0;
            ISheet sheet = XslHelper.OpenSheet(FilePath, true, ref StartRow, ref StartCell, ref Mistakes, Type);
            if (sheet == null)
            {
                if (Error.ContainsKey("表格格式内容"))
                {
                    Error["表格格式内容"].Add("提交的表而过无法检索，请核对格式");
                }
                else
                {
                    Error.Add("表格格式内容", new List<string> { "提交的表格无法检索，请核对格式" });
                }
                return false;
            }
            StartRow++;
            int Max = sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++) {
                var row = sheet.GetRow(i);
                if (row == null)
                    break;
                var value = row.Cells[StartCell + 3].GetValue().Replace(" ","").ToString().Trim();
                var county = row.Cells[StartCell + 2].GetValue().Replace(" ","").ToString().Trim();
                var Name = row.Cells[StartCell + 4].GetValue().Replace(" ","").ToString().Trim();
                var key = Name.ToUpper() + '-' + county.ToUpper() + '-' + value.ToUpper();
                if (string.IsNullOrEmpty(value) && string.IsNullOrEmpty(county) && string.IsNullOrEmpty(Name))
                    continue;
                if (Error.ContainsKey(key))
                    continue;
                var SeProject = new SeProject {
                    IsHasDoubt = row.Cells[StartCell + 5].GetValue() == "是",
                    IsApplyDelete = row.Cells[StartCell + 6].GetValue() == "是",
                    IsHasError = row.Cells[StartCell + 7].GetValue() == "是",
                    IsPacket = row.Cells[StartCell + 8].GetValue() == "是",
                    IsDescrease = row.Cells[StartCell + 9].GetValue() == "是",
                    IsRelieve = row.Cells[StartCell + 10].GetValue() == "是"
                };
                PlanData.Add(new pProject
                {
                    Key=value,
                    Name=Name,
                    County=county,
                    IsApplyDelete=SeProject.IsApplyDelete,
                    IsHasError=SeProject.IsHasError,
                    IsRight=!(SeProject.IsHasDoubt||SeProject.IsPacket||SeProject.IsDescrease||SeProject.IsRelieve)
                });
            }
            Fit();
                return true;
        }

        public void  Fit() {
            foreach (var item in PlanIDS.Keys) {
                if (PlanIDS[item] > 0) {
                    if (Error.ContainsKey(item))
                    {
                        List<string> Buffer = Error[item];
                        Buffer.Add("规则0002（一致性）：复核确认未验收项目清单中存在该项目，但是不存在在本表中");
                        Error[item] = Buffer;
                    }
                    else {
                        Error.Add(item, new List<string>() { "规则0002（一致性）：复核确认未验收项目清单中存在该项目，但是不存在在本表中" });
                    }
                }
            }
        }

        public bool GetProject(string FilePath,ref string Misatkes, SecondReportType Type) {
            int StartRow = 0, StartCell = 0;
            ISheet sheet = XslHelper.OpenSheet(FilePath, true, ref StartRow, ref StartCell, ref Misatkes, Type);
            if (sheet == null) {
                if (Error.ContainsKey("表格格式内容"))
                {
                    Error["表格格式内容"].Add("提交的表而过无法检索，请核对格式");
                }
                else {
                    Error.Add("表格格式内容", new List<string> { "提交的表格无法检索，请核对格式" });
                }
                return false;
            }
            StartRow++;
            int Max = sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++) {
                var row = sheet.GetRow(i);
                if (row == null)
                    break;
                var value = row.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                if (!value.VerificationID())
                    continue;
                if (Error.ContainsKey(value))
                    continue;
                if (!Data.ContainsKey(value)) {
                    Data.Add(value, new SeProject
                    {
                        IsHasDoubt = row.Cells[StartCell + 5].GetValue() == "是",
                        IsApplyDelete = row.Cells[StartCell + 6].GetValue() == "是",
                        IsHasError = row.Cells[StartCell + 7].GetValue() == "是",
                        IsPacket=row.Cells[StartCell+8].GetValue()=="是",
                        IsDescrease = row.Cells[StartCell + 9].GetValue() == "是",
                        IsRelieve = row.Cells[StartCell + 10].GetValue() == "是"
                    });
                }
            }
              
            return true;
        }
    }
}