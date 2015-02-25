using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public static class GetEight
    {
        public static List<string> Operation(string FilePath) {
            IWorkbook workbook = null;
            using (var fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read)) {
                workbook = WorkbookFactory.Create(fs);
            }
            if (workbook == null) {
                return null;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            if (sheet == null) {
                return null;
            }
            List<string> ProjectID = new List<string>();
            int Max = sheet.LastRowNum;
            int[] Lines={5,6,7};
            IRow row = null;
            ICell cell = null;
            for (var i = 5; i <= Max; i++) {
                row = sheet.GetRow(i);
                if (row == null) {
                    break;
                }
                var value = row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
                if (string.IsNullOrEmpty(value)) {
                    continue;
                }
                double[] data=new double[3];
                int j=0;
                foreach(int Index in Lines){
                    cell=row.GetCell(Index,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if (cell.CellType == CellType.Formula || cell.CellType == CellType.Numeric)
                    {
                        try
                        {
                            data[j] = cell.NumericCellValue;
                        }
                        catch
                        {
                            data[j] = 0.0;
                        }
                    }
                    else {
                        double.TryParse(cell.ToString(), out data[j]);
                    }
                    j++;
                }
                if (Math.Abs(data[0] - data[1]) < 0.0001 || Math.Abs(data[0] - data[2]) < 0.0001) {
                    ProjectID.Add(value);
                }
               
                
            }
            return ProjectID;
        }

        public static bool Save(List<string> Data,string FilePath) {
            throw new NotImplementedException();
            
        }

    }
}