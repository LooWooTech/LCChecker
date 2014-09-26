using NPOI.SS.UserModel;

namespace LCChecker.Rules
{
    internal class Format:IRowRule
    {
        public int ColumnIndex { get; set; }
        public string form { get; set; }
        public string Name {
            get {
                return string.Format("第{0}栏的格式为：{1} 或者1到15", ColumnIndex + 1, form);
            }
        }
        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            var strs = value.Split('.');
            try
            {
                int k = int.Parse(strs[0]);
            }
            catch {
                return false;
            }
            return true;
        }
    }
}