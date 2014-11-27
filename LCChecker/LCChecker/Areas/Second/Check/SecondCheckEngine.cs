using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public class SecondCheckEngine
    {
        public List<RuleInfo> rules = new List<RuleInfo>();
        public Dictionary<string, List<string>> Error = new Dictionary<string, List<string>>();
        public Dictionary<string, string> Warning = new Dictionary<string, string>();
        public List<string> IDS = new List<string>();
        public Dictionary<string, bool> Whether = new Dictionary<string, bool>();


        public Dictionary<string, List<string>> GetError() {
            return Error;
        }
        public List<string> GetIDS() {
            return IDS;
        }

        public bool Check(string FilePath, ref string Mistakes, SecondReportType Type) {
            return CheckEngine(FilePath, ref Mistakes, Type);
        }
       
        public bool CheckEngine(string FilePath, ref string Mistakes, SecondReportType Type) {
            int StartRow = 0, StartCell = 0;
            ISheet sheet = XslHelper.OpenSheet(FilePath, true, ref StartRow, ref StartCell, ref Mistakes, Type);
            if (sheet == null) {
                if (Error.ContainsKey("表格格式内容"))
                {
                    Error["表格格式内容"].Add("提交的表格无法检索，请核对格式");
                }
                else {
                    Error.Add("表格格式内容", new List<string> { "提交的表格无法检索，请核对格式" });
                }
                return false;
            }
            StartRow++;
            int Max=sheet.LastRowNum;
            for (var i = StartRow; i <= Max; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    break;
                List<string> ErrorRow = new List<string>();
                var value = row.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;
                if (!value.VerificationID())
                    continue;
                if (IDS.Contains(value))
                {
                    if (Error.ContainsKey(value))
                    {
                        Error[value].Add("表格中存在相同项目编号");
                    }
                    else {
                        Error.Add(value, new List<string> { "表格中存在相同项目编号" });
                    }
                    continue;
                }
                IDS.Add(value);
                if (Whether.ContainsKey(value))
                {
                    if (!Whether[value]) {
                        ErrorRow.Add("规则000：与重点复核确认项目以外所有报部备案项目复核确认总表不符");
                        //Warning[value] = "规则000：与重点复核确认项目以外所有报部备案项目复核确认总表不符";
                    }
                    foreach (var item in rules) {
                        if (!item.Rule.Check(row, StartCell)) {
                            ErrorRow.Add(item.Rule.Name);
                        }
                    }
                    if (ErrorRow.Count() != 0) {
                        if (Error.ContainsKey(value))
                        {
                            Error[value] = ErrorRow;
                        }
                        else {
                            Error.Add(value, ErrorRow);
                        }
                    }

                    Whether.Remove(value);
                }
                else {
                    ErrorRow.Add("规则000：重点复核确认总表中不包括该项目");
                    if (!Error.ContainsKey(value))
                    {
                        Error.Add(value, ErrorRow);
                    }
                }

            }
            foreach (var item in Whether.Keys) {
                if (Whether[item]) {
                    if (Error.ContainsKey(item))
                    {
                        Error[item].Add("规则000：项目存在复核确认验收清单中，但是不存在本表中");
                    }
                    else {
                        Error.Add(item, new List<string> { "规则000：项目存在复核确认验收清单中，但是不存在本表中" });
                    }
                    
                    //if (Warning.ContainsKey(item))
                    //{
                    //    Warning[item] += "规则000：项目存在复核确认验收清单中，但是不存在本表中";
                    //}
                    //else {
                    //    Warning.Add(item, "规则000：项目存在复核确认验收清单中，但是不存在本表中");
                    //}
                }
            }   
            return true;
        }

        
    }
}