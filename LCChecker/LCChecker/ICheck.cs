using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LCChecker
{
    public interface ICheck
    {
        Dictionary<string, List<string>> GetError();
        bool Check(string filePath, ref string mistakes);
    }
}
