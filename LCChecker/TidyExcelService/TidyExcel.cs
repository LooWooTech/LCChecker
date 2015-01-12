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
            string error = string.Empty;
            string TemplatePath = string.Empty;
            if (!GetPath(ref TemplatePath, ref error))
            {
                var Loggerror = log4net.LogManager.GetLogger("logerror");
                Loggerror.ErrorFormat("读取路径时发生错误：{0}", error);
                return;
            }
            try
            {
                using (var coon = new MySqlConnection(ConfigurationManager.ConnectionStrings["TIDY"].ConnectionString))
                {
                    coon.Open();
                    int id = 0;
                    int CityID = 0;
                    string SavePath = "";
                    int Type = 0;
                    bool Flag = false;
                    using (var cmd = coon.CreateCommand())
                    {
                        cmd.CommandText = "SELECT ID,CityID,SavePath,Type,IsPlan FROM `uploadfiles` WHERE Type BETWEEN 20 And 30  AND State=1 AND Census=0 ORDER BY ID";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read() == false) return;

                            id = Convert.ToInt32(reader[0]);
                            CityID = Convert.ToInt32(reader[1]);
                            SavePath = reader[2].ToString();
                            Type = Convert.ToInt32(reader[3]);
                            Flag = Convert.ToBoolean(reader[4]);
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
                            if (Flag)
                            {
                                cmd.CommandText = string.Format("SELECT SavePath from uploadfiles WHERE CreateTime=(SELECT MAX(CreateTime) FROM `uploadfiles` WHERE CityID={0} ANd type={1} AND State=1 AND IsPlan=1)", (int)item, Type);
                            }
                            else {
                                cmd.CommandText = string.Format("SELECT SavePath from uploadfiles WHERE CreateTime=(SELECT MAX(CreateTime) FROM `uploadfiles` WHERE CityID={0} ANd type={1} AND State=1 AND IsPlan=0)", (int)item, Type);
                            }
                            
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.Read() == false) continue; 
                                string SavePaths = Path.Combine(TemplatePath, reader[0].ToString());
                                FilesPath.Add(SavePaths);
                            }
                        }
                    }
                    SecondReportType NewType = (SecondReportType)(Type - 20);
                    //if (NewType == SecondReportType.附表8)
                    //{
                    //    TidyEight(FilesPath);
                    //}
                    //else {
                    //    Tidy(FilesPath, NewType,Flag);
                    //}

                    if (NewType != SecondReportType.附表8) {
                       // TemplatePath = Path.Combine(TemplatePath, "templates");
                        Tidy(FilesPath, NewType, Flag,TemplatePath);
                    }

                    


                    using (var cmd2 = coon.CreateCommand()) {
                        cmd2.CommandText = string.Format("UPDATE uploadfiles SET Census=1 WHERE ID={0}", id);
                        cmd2.ExecuteNonQuery();
                    }






                }
            }
            catch (Exception er) {
                var Loggerror = log4net.LogManager.GetLogger("logerror");
                Loggerror.ErrorFormat("数据库操作时发生错误：{0}",er);
            }
        }


        public static void Tidy(List<string> FilePaths,SecondReportType Type,bool IsPlan,string Template) {
            string TemplatePath = Path.Combine(Template,"templates/Second", Type.ToString() + ".xls");
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
                case SecondReportType.附表1: StartNumber = 6; Lines=11; xoffset=2;break;
                case SecondReportType.附表2: StartNumber = 6; Lines = 10; xoffset=2;break;
                case SecondReportType.附表3: StartNumber = 6; Lines = 10; xoffset=3;break;
                case SecondReportType.附表4: StartNumber = 5; Lines = 9; xoffset=3;break;
                case SecondReportType.附表6: StartNumber = 5; Lines = 9; xoffset=4;break;
                case SecondReportType.附表7: StartNumber = 5; Lines = 11; xoffset=3;break;

                case SecondReportType.附表8: StartNumber = 6; Lines = 25; xoffset = 4; return;
                
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
                for (var i = StartNumber; i <= Max; i=i+AddLine)
                {
                    IRow MacRow = MacSheet.GetRow(i);
                    if (MacRow == null)
                        break;
                    var value = MacRow.GetCell(StartCell + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                    if (!IsPlan) {
                        if (string.IsNullOrEmpty(value))
                            continue;
                        if (!value.VerificationID())
                            continue;
                    }
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
                        Copy(ref rowOne, ref MacRow, 1, StartCell + 1, Lines - 1);

                        #region
                        //for (var j=1;j<Lines;j++){
                        //    var cell=rowOne.GetCell(j);
                        //    var maccell=MacRow.GetCell(j);
                        //    if(maccell==null){
                        //        cell.SetCellValue("");
                        //        continue;
                        //    }
                        //    switch(maccell.CellType){
                        //        case CellType.Boolean:cell.SetCellValue(maccell.BooleanCellValue);break;
                        //        case CellType.Numeric:cell.SetCellValue(maccell.NumericCellValue);break;
                        //        case CellType.String:cell.SetCellValue(maccell.StringCellValue);break;
                        //        case CellType.Formula:
                        //            double data=.0;
                        //            try{
                        //                data=maccell.NumericCellValue;
                        //            }catch{
                        //                data=.0;
                        //            }
                        //            cell.SetCellValue(data);break;
                        //        case CellType.Blank:cell.SetCellValue("");break;
                        //        default:cell.SetCellValue(maccell.ToString().Trim());break;
                        //    }

                        //}
                        #endregion
                        
                        int m=1;
                         for(var j=StartNumber-2;j<StartNumber;j++){
                             rowOne=sheet.GetRow(j);
                             MacRow=MacSheet.GetRow(i+m);
                             m++;
                             Copy(ref rowOne, ref MacRow, 6, StartCell + 6, Lines - 6);
                             #region
                             //for (var k=6;k<Lines;k++){
                             //   var cell=rowOne.GetCell(k);
                             //    var maccell=MacRow.GetCell(k);
                             //    if(maccell==null)
                             //    {
                             //        cell.SetCellValue("");
                             //        continue;
                             //    }
                             //    switch(maccell.CellType){
                             //        case CellType.Numeric:cell.SetCellValue(maccell.NumericCellValue);break;
                             //        case CellType.String:cell.SetCellValue(maccell.StringCellValue);break;
                             //        case CellType.Boolean:cell.SetCellValue(maccell.BooleanCellValue);break;
                             //        case CellType.Formula:
                             //            double data=0.0;
                             //            try{
                             //               data=maccell.NumericCellValue;
                             //            }catch{
                             //               data=0.0;
                             //            }
                             //            cell.SetCellValue(data);break;
                             //       case CellType.Blank: cell.SetCellValue(""); break;
                             //       case CellType.Unknown: cell.SetCellValue(""); break;
                             //       case CellType.Error: cell.SetCellValue(""); break;
                             //       default: cell.SetCellValue(maccell.ToString().Trim()); break;
                             //    }
                             //}
                             #endregion
                         }
                    }
                    else {
                        IRow row = sheet.GetRow(StartNumber);
                        if (row == null) {
                            row=sheet.CreateRow(StartNumber);
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
            if (IsPlan)
            {
                TemplatePath = Path.Combine(Template, "App_Data", Type.ToString() + "-未验收总表.xls");
            }
            else {
                TemplatePath = Path.Combine(Template, "App_Data", Type.ToString() + "-验收总表.xls");
            }
            
            using (var fs = new FileStream(TemplatePath, FileMode.OpenOrCreate, FileAccess.Write)) {
                workbook.Write(fs);
                fs.Flush();
            }
        }


        public static void TidyEight(List<string> FilePaths)
        {
            //打开模板文件
            #region
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
            #endregion
            int StartNumber = 6;
            int SerialNumber = 1;
            IRow TemplateRow=Sheet.GetRow(StartNumber);
            //遍历所有市的附表8
            #region
            foreach (var file in FilePaths) {
                
                IWorkbook Macbook = null;
                //打开某个市上传的附表8
                #region
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
                #endregion

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
                    int SpanR1 = 0,SpanC1=0;
                    //
                    #region
                    if (XslHelper.isMergeCell(macsheet, i + 1, StartCell + 3, out SpanR1, out SpanC1))
                    {

                        //当读取到的表格一个项目占据了SpanR1行的单元格，那么表格往下移动SpanR1行，
                        //保存数据的表格 创建从StartNumber行开始一共创建SpanR1行25列  
                        //合并前9列单元格  一个单元格占据SpanR1行
                        #region
                        Sheet.ShiftRows(Sheet.LastRowNum - 4, Sheet.LastRowNum, SpanR1, true, false);

                        for (var j = StartNumber; j < (StartNumber + SpanR1); j++)
                        {
                            var Bufferrow1 = Sheet.GetRow(j);
                            if (Bufferrow1 == null)
                            {
                                Bufferrow1 = Sheet.CreateRow(j);
                                if (TemplateRow.RowStyle != null)
                                {
                                    Bufferrow1.RowStyle = TemplateRow.RowStyle;
                                }
                            }
                            for (var k = 0; k < 25; k++)
                            {
                                var cell = Bufferrow1.GetCell(k);
                                if (cell == null)
                                {
                                    cell = Bufferrow1.CreateCell(k, TemplateRow.GetCell(k).CellType);
                                    cell.CellStyle = TemplateRow.GetCell(k).CellStyle;
                                }
                            }
                        }
                        for (var j = 0; j < 9; j++)
                        {
                            Sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(StartNumber, StartNumber + SpanR1 - 1, j, j));
                        }
                        #endregion


                        //拷贝填写已挂钩补充耕地项目
                        //读取源文件前9列的数据 macrow  从startCell列到startCell+9列获取
                        #region
                        IRow row = Sheet.GetRow(StartNumber);
                        row.GetCell(0).SetCellValue(SerialNumber++);

                        //拷贝数据
                        Copy(ref row, ref macrow, 1, StartCell + 1, 8);

                        #endregion


                        //对应建设用地项目情况
                        //开始从对应建设用地项目遍历
                        //判断从i+offset行 startCell+12列单元格  是否合并
                        #region
                        int FlagR = SpanR1;
                        int Offset = 0;
                        int Serial1 = 1;
                        while (FlagR > 0)
                        {
                            int SpanR2 = 0, SpanC2 = 0;

                            if (XslHelper.isMergeCell(macsheet, i + Offset + 1, StartCell + 12, out SpanR2, out SpanC2))
                            {
                                macrow = macsheet.GetRow(i + Offset);

                                //获取对应建设用地项目的项目编号并且判断
                                //合并对应建设用地项目单元格
                                #region
                                var key2 = macrow.GetCell(StartCell + 12, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
                                if (string.IsNullOrEmpty(key2))
                                    continue;
                                if (!key2.VerificationID())
                                {
                                    continue;
                                }

                                //
                                for (var j = 9; j < 15; j++)
                                {
                                    Sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(StartNumber + Offset, StartNumber + SpanR2 - 1 + Offset, j, j));
                                }
                                //将源文件中的数据拷贝到总表中
                                var Bufferrow2 = Sheet.GetRow(StartNumber + Offset);
                                //Bufferrow2.GetCell(Serial1++);
                                Bufferrow2.GetCell(9).SetCellValue(Serial1++);
                                Copy(ref Bufferrow2, ref macrow, 10, StartCell + 10, 5);

                                #endregion

                                int serial2 = 1;
                                for (var j = 0; j < SpanR2; j++)
                                {
                                    macrow = macsheet.GetRow(j + i + Offset);
                                    if (macrow == null)
                                        continue;
                                    row = Sheet.GetRow(StartNumber + Offset + j);
                                    row.GetCell(15).SetCellValue(serial2++);
                                    //
                                    Copy(ref row, ref macrow, 16, StartCell + 16, 9);
                                }
                            }
                            else
                            {
                                row = Sheet.GetRow(StartNumber + Offset);
                                macrow = macsheet.GetRow(i + Offset);
                                row.GetCell(9).SetCellValue(Serial1++);
                                Copy(ref row, ref macrow, 10, StartCell + 10, 15);
                            }
                            FlagR -= SpanR2;
                            Offset += SpanR2;
                        }
                        #endregion



                    }
                    else {
                        var row = Sheet.GetRow(StartNumber);
                        Copy(ref row, ref macrow, 0, StartCell, 25);
                        row.GetCell(0).SetCellValue(SerialNumber++);
                    }
                    #endregion

                    i = i + SpanR1 - 1;
                    StartNumber += SpanR1;
                }

            }
            #endregion

            TemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../LCChecker/App_Data", SecondReportType.附表8.ToString() + "-总表.xls");
            using (var fs = new FileStream(TemplatePath, FileMode.OpenOrCreate, FileAccess.Write)) {
                workbook.Write(fs);
                fs.Flush();
            }


        }

        /// <summary>
        /// 拷贝数据
        /// </summary>
        /// <param name="Row">拷贝的行</param>
        /// <param name="SourceRow">被拷贝的行</param>
        /// <param name="StartCell1">拷贝行开始的位置</param>
        /// <param name="StartCell2">被拷贝行开始的位置</param>
        /// <param name="Rank">拷贝多少列</param>

        public static void Copy(ref IRow Row, ref IRow SourceRow,int StartCell1,int StartCell2,int Rank) {
            for (var i = 0; i < Rank; i++) {
                var cell = Row.GetCell(StartCell1+i);
                var SourceCell = SourceRow.GetCell(StartCell2+i);
                if (SourceCell == null) {
                    cell.SetCellValue("");
                    continue;
                }
                switch (SourceCell.CellType) {
                    case CellType.Boolean: cell.SetCellValue(SourceCell.BooleanCellValue); break;
                    case CellType.Numeric: cell.SetCellValue(SourceCell.NumericCellValue); break;
                    case CellType.String: cell.SetCellValue(SourceCell.StringCellValue); break;
                    case CellType.Formula:
                        double data = 0.0;
                        try
                        {
                            data = SourceCell.NumericCellValue;
                        }
                        catch { 
                        
                        }
                        cell.SetCellValue(data);break;
                    default: cell.SetCellValue(SourceCell.ToString().Trim()); break;
                }
            }
        }


        public static bool GetPath(ref string FilePath,ref string Error) {
            var BaseFolder = ConfigurationManager.AppSettings["BaseFolder"];
            DirectoryInfo dir = new DirectoryInfo(BaseFolder);
            if (!dir.Exists) {
                return false;
            }
            try {
                string file = Path.Combine(BaseFolder, "templatePath.txt");
                FileStream fs = new FileStream(file, FileMode.Open);
                StreamReader sr = new StreamReader(fs);
                FilePath = sr.ReadLine();
                sr.Close();
            }catch(IOException e){
                Error = e.ToString();
                return false;
            }
            return true;
        }


    }
}
