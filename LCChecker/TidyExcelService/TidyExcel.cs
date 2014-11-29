using LCChecker;
using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using MySql.Data.MySqlClient;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace TidyExcelService
{
    public class TidyExcel
    {

        public static void Process() {
            try
            {
                using (var coon = new MySqlConnection(ConfigurationManager.ConnectionStrings["TIDY"].ConnectionString))
                {
                    coon.Open();
                    int id = 0;
                    int CityID = 0;
                    string SavePath = "";
                    int Type = 0;
                    using (var cmd = coon.CreateCommand())
                    {
                        cmd.CommandText = "SELECT ID,CityID,SavePath,Type FROM `uploadfiles` WHERE Type BETWEEN 4 And 9  AND State=1 AND Census=0 ORDER BY ID";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read() == false) return;

                            id = Convert.ToInt32(reader[0]);
                            CityID = Convert.ToInt32(reader[1]);
                            SavePath = reader[2].ToString();
                            Type = Convert.ToInt32(reader[3]);
                            reader.Close();
                        }
                    }

                    List<string> FilesPath = new List<string>();
                    foreach (City item in Enum.GetValues(typeof(City)))
                    {
                        if (item == City.浙江省)
                        {
                            continue;
                        }
                        using (var cmd = coon.CreateCommand())
                        {
                            cmd.CommandText = string.Format("SELECT SavePath from uploadfiles WHERE CreateTime=(SELECT MAX(CreateTime) FROM `uploadfiles` WHERE CityID={0} ANd type={1} AND State=1)", (int)item, Type);
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.Read() == false) return;
                                string SavePaths = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../LCChecker/", reader[0].ToString());
                                FilesPath.Add(SavePaths);
                            }
                        }
                    }







                }
            }
            catch (Exception er) {
                var Loggerror = log4net.LogManager.GetLogger("logerror");
                Loggerror.ErrorFormat("数据库操作时发生错误：{0}",er);
            }
        }


        public static void Tidy(List<string> FilePaths,SecondReportType Type) {
            string TemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../LCChecker/Templates/Second", Type.ToString() + ".xls");
            IWorkbook workbook = null;
            try
            {
                using (var fs = new FileStream(TemplatePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch {
                return;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            int id = 1;
            int StartNumber = 0, xoffset = 0, Lines = 0;
            int AddLine = 1;
            int[] Merge = { 0, 1, 2, 3, 4, 5, 22 };
            switch (Type) {
                case SecondReportType.附表1: StartNumber = 6; Lines=11; xoffset=3;break;
                case SecondReportType.附表2: StartNumber = 6; Lines = 10; xoffset=3;break;
                case SecondReportType.附表3: StartNumber = 6; Lines = 10; xoffset=4;break;
                case SecondReportType.附表4: StartNumber = 6; Lines = 9; xoffset=4;break;
                case SecondReportType.附表6: StartNumber = 5; Lines = 9; xoffset=5;break;
                case SecondReportType.附表7: StartNumber = 5; Lines = 11; xoffset=4;break;
                case SecondReportType.附表8: StartNumber = 6; Lines = 25; xoffset=4;break;
                case SecondReportType.附表9: StartNumber = 6; Lines = 23; xoffset = 3; AddLine = 3; break;
                default: break;
            }

            IRow[] TemplateRow = new IRow[AddLine];
            for (var k = 0; k < AddLine; k++)
            {
                TemplateRow[k] = sheet.GetRow(StartNumber + k);
            }
                //sheet.GetRow(StartNumber);

            foreach (var file in FilePaths) {
                IWorkbook Macbook = null;
                try
                {
                    using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        Macbook = WorkbookFactory.Create(fs);
                    }
                }
                catch {
                    continue;
                }
                ISheet MacSheet = Macbook.GetSheetAt(0);
                if (MacSheet == null)
                    continue;
                int StartRow = 0, StartCell = 0;
                if (!XslHelper.FindHeader(MacSheet, ref StartRow, ref StartCell, Type)) {
                    continue;
                }
                StartRow++;
                int Max=MacSheet.LastRowNum;
                for (var i = 0; i <= Max; i=i+AddLine)
                {
                    IRow MacRow = MacSheet.GetRow(i);
                    if (MacRow == null)
                        break;
                    var value = MacRow.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                    if (string.IsNullOrEmpty(value))
                        continue;
                    if (!value.VerificationID())
                        continue;
                    sheet.ShiftRows(sheet.LastRowNum - xoffset, sheet.LastRowNum, AddLine, true, false);
                    if (Type == SecondReportType.附表9)
                    {
                        foreach (var Trow in TemplateRow) {
                            IRow row= sheet.GetRow(StartNumber);
                            if (row == null) {
                                row = sheet.CreateRow(StartNumber);
                                row.RowStyle = Trow.RowStyle;
                            }
                            StartNumber++;
                            for (var n = 0; n < 23; n++) {
                                var maccell = Trow.GetCell(n);
                                var cell = row.GetCell(n);
                                if (cell == null) {
                                    cell = row.CreateCell(n, maccell.CellType);
                                    cell.CellStyle = maccell.CellStyle;
                                }
                            }
                        }
                        foreach (var Mcell in Merge) {
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(StartNumber - 3, StartNumber - 1, Mcell, Mcell));
                        }
                        IRow rowOne=sheet.GetRow(StartNumber-3);
                        rowOne.GetCell(0).SetCellValue(id++);
                        for(var j=1;j<Lines;j++){
                            var cell=rowOne.GetCell(j);
                            var maccell=MacRow.GetCell(j);
                            if(maccell==null){
                                cell.SetCellValue("");
                                continue;
                            }
                            switch(maccell.CellType){
                                case CellType.Boolean:cell.SetCellValue(maccell.BooleanCellValue);break;
                                case CellType.Numeric:cell.SetCellValue(maccell.NumericCellValue);break;
                                case CellType.String:cell.SetCellValue(maccell.StringCellValue);break;
                                case CellType.Formula:
                                    double data=.0;
                                    try{
                                        data=maccell.NumericCellValue;
                                    }catch{
                                        data=.0;
                                    }
                                    cell.SetCellValue(data);break;
                                case CellType.Blank:cell.SetCellValue("");break;
                                default:cell.SetCellValue(maccell.ToString().Trim());break;
                            }
                           
                         }
                         int m=1;
                         for(var j=StartNumber-2;j<StartNumber;j++){
                             rowOne=sheet.GetRow(j);
                             MacRow=MacSheet.GetRow(i+m);
                             m++;
                             for(var k=6;k<Lines;k++){
                                var cell=rowOne.GetCell(k);
                                 var maccell=MacRow.GetCell(k);
                                 if(maccell==null)
                                 {
                                     cell.SetCellValue("");
                                     continue;
                                 }
                                 switch(maccell.CellType){
                                     case CellType.Numeric:cell.SetCellValue(maccell.NumericCellValue);break;
                                     case CellType.String:cell.SetCellValue(maccell.StringCellValue);break;
                                     case CellType.Boolean:cell.SetCellValue(maccell.BooleanCellValue);break;
                                     case CellType.Formula:
                                         double data=0.0;
                                         try{
                                            data=maccell.NumericCellValue;
                                         }catch{
                                            data=0.0;
                                         }
                                         cell.SetCellValue(data);break;
                                    case CellType.Blank: cell.SetCellValue(""); break;
                                    case CellType.Unknown: cell.SetCellValue(""); break;
                                    case CellType.Error: cell.SetCellValue(""); break;
                                    default: cell.SetCellValue(maccell.ToString().Trim()); break;
                                 }
                             }
                        }
                    }
                    else {
                        IRow row = sheet.GetRow(StartNumber);
                        if (row == null) {
                            row.Sheet.CreateRow(StartNumber);
                            if (TemplateRow[0].RowStyle != null) {
                                row.RowStyle = TemplateRow[0].RowStyle;
                            }
                        }
                        StartNumber++;
                        var cell = row.GetCell(0);
                        if (cell == null) {
                            cell = row.CreateCell(0, TemplateRow[0].GetCell(0).CellType);
                            cell.CellStyle = TemplateRow[0].GetCell(0).CellStyle;
                        }
                        cell.SetCellValue(id++);
                        for (var j = 1; j < Lines; j++) {
                            var maccell = MacRow.GetCell(j + StartCell);
                            if (maccell == null) {
                                continue;
                            }
                            cell = row.GetCell(j);
                            if (cell == null) {
                                cell = row.CreateCell(j, TemplateRow[0].GetCell(j).CellType);
                                cell.CellStyle = TemplateRow[0].GetCell(j).CellStyle;
                            }
                            switch (maccell.CellType) {
                                case CellType.Boolean: cell.SetCellValue(maccell.BooleanCellValue); break;
                                case CellType.Numeric: cell.SetCellValue(maccell.NumericCellValue); break;
                                case CellType.String: cell.SetCellValue(maccell.StringCellValue); break;
                                case CellType.Formula:
                                    double data = 0.0;
                                    try
                                    {
                                        data = maccell.NumericCellValue;
                                    }
                                    catch {
                                        data = 0.0;
                                    }
                                    cell.SetCellValue(data);
                                    break;
                                default: cell.SetCellValue(maccell.ToString()); break;
                            }
                        }
                    }


                }
            }
            TemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../LCChecker/App_Data", Type.ToString() + "-总表.xls");
            using (var fs = new FileStream(TemplatePath, FileMode.OpenOrCreate, FileAccess.Write)) {
                workbook.Write(fs);
                fs.Flush();
            }
        }


        public static void TidyEight(List<string> FilePaths) {
            string TemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../LCChecker/Templates/Second/附表8.xls");
            IWorkbook workbook = null;
            try
            {
                using (var fs = new FileStream(TemplatePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
            }
            catch(Exception er) {
                var Logerror = log4net.LogManager.GetLogger("logerror");
                Logerror.ErrorFormat("读取附表8模板失败，错误：{0}", er);
                return;
            }
            ISheet Sheet = workbook.GetSheetAt(0);
            if (Sheet == null)
                return;
            int StartNumber = 6;
            IRow TemplateRow=Sheet.GetRow(StartNumber);
            foreach (var file in FilePaths) {
                IWorkbook Macbook = null;
                try
                {
                    using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        Macbook = WorkbookFactory.Create(fs);
                    }
                }
                catch (Exception er) {
                    var Logerror = log4net.LogManager.GetLogger("logerror");
                    Logerror.ErrorFormat("在整理附表8的时候读取文件{0}失败，错误：{1}",file, er);
                    continue;
                }
                ISheet macsheet = Macbook.GetSheetAt(0);
                if (macsheet == null)
                    continue;
                int StartRow = 0, StartCell = 0;
                if (!XslHelper.FindHeader(macsheet, ref StartRow, ref StartCell, SecondReportType.附表8)) {
                    continue;
                }
                StartRow++;
                int Max=macsheet.LastRowNum;
                for (var i = StartRow; i <= Max; i++) {
                    IRow macrow = macsheet.GetRow(i);
                    if (macrow == null)
                        break;
                    var value = macrow.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                    if (string.IsNullOrEmpty(value))
                        continue;
                    if (!value.VerificationID())
                        continue;
                    int SpanR = 0,SpanC=0;
                    if (XslHelper.isMergeCell(macsheet, i, StartCell + 3, out SpanR, out SpanC)) { 
                        
                    }
                }

            }
        }


    }
}
