﻿using LCChecker.Areas.Second.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LCChecker.Areas.Second
{
    public interface ISeCheck
    {
        Dictionary<string, List<string>> GetError();
        List<string> GetIDS();
        Dictionary<string, SeProject> GetSeProject();
        Dictionary<string, SeLand> GetPaddy();
        Dictionary<string, SeLand> GetDry();
        bool Check(string FilePath, ref string Mistakes, SecondReportType Type);
    }
}