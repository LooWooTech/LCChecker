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
        public List<string> IDS = new List<string>();
        public Dictionary<string, bool> Whether = new Dictionary<string, bool>();
        public bool Check(string FilePath, ref string Mistakes, SecondReportType Type, List<SecondProject> Data) {
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


            }
                return true;
        }

        
    }
}