﻿using LCChecker.Areas.Second.Models;
using LCChecker.Areas.Second.Rules;
using LCChecker.Models;
using LCChecker.Rules;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace LCChecker.Areas.Second
{
    public class CheckOne:SecondCheckEngine,ISeCheck
    {
       

        public CheckOne(List<SecondProject> projects) {
            Whether = projects.ToDictionary(e => e.ID, e => true);
            Dictionary<string,SecondProject> Team=projects.ToDictionary(e=>e.ID,e=>e);
            var list = new List<IRowRule>();
            list.Add(new OnlySecondProject() { ColumnIndex = 3, AreaIndex = 0, NewAreaIndex = 0, Projects = Team,ID="2101", Values = new[] { "项目名称", "市", "县" } });
            list.Add(new CellRangeRowRule() { ColumnIndex = 5, Values = new[] { "是", "否" }, ID = "2102" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是", "否" }, ID = "2103" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "是", "否" }, ID = "2104" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" }, ID = "2105" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 9, Values = new[] { "是", "否" }, ID = "2106" });
            list.Add(new CellRangeRowRule() { ColumnIndex = 10, Values = new[] { "是", "否" }, ID = "2107" });
            list.Add(new ConditionalRowRule()
            {
                Condition = new CellRangeRowRule() { ColumnIndex = 6, Values = new[] { "是" } },
                Rule = new CellRangeRowRule() { ColumnIndex = 7, Values = new[] { "否" } }
            });
            foreach (var item in list) {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }


        public new bool Check(string FilePath, ref string Mistakes, SecondReportType Type) {
            if (!CheckEngine(FilePath, ref Mistakes, Type)) {
                return false; 
            }
            if (!GetProject(FilePath, ref Mistakes, Type)) {
                return false;
            }
            return true;
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
                if (!Data.ContainsKey(value)) {
                    Data.Add(value, new SeProject
                    {
                        IsHasDoubt = row.Cells[StartCell + 5].GetValue() == "是",
                        IsApplyDelete = row.Cells[StartCell + 6].GetValue() == "是",
                        IsHasError = row.Cells[StartCell + 7].GetValue() == "是",
                        IsDescrease = row.Cells[StartCell + 9].GetValue() == "是",
                        IsRelieve = row.Cells[StartCell + 10].GetValue() == "是"
                    });
                }
            }
              
            return true;
        }
    }
}