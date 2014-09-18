using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LCChecker
{
    public interface ICheck
    {
        Dictionary<string, List<string>> GetError();

        Dictionary<string, string> GetWarning();
        bool Check(string filePath, ref string mistakes,ReportType Type,List<Project> Data,bool flag);
    }
}
