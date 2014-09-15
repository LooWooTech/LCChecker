using NPOI.SS.UserModel;

namespace LCChecker.Rules
{
    internal class Format:IRowRule
    {
        public int ColumnIndex { get; set; }
        public string form { get; set; }
        public string Name {
            get {
                return string.Format("第{0}栏的格式为：{1}", ColumnIndex + 1, form);
            }
        }
        public bool Check(NPOI.SS.UserModel.IRow row, int xoffset = 0)
        {
            var value = row.GetCell(ColumnIndex + xoffset, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString().Trim();
            var strs = value.Split('.');
            int count=strs.Length;
            if (count==1||strs[count-1]=="")//这样做 可以考虑到12. 假如只是  Contains（）的话 
            {
                return false;
            }
            return true;
        }
    }
}