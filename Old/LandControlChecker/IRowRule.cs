using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LooWoo.Land.LandControlChecker
{
    public interface IRowRule
    {
        string Name { get; }
        bool Check(IRow row, int xoffset = 0);
    }
}
