using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LCChecker
{
    public interface IRowRule
    {
        string Name { get; }
        bool Check(IRow row, int xoffset = 0);
    }
}
