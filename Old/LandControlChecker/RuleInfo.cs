using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LooWoo.Land.LandControlChecker
{
    public class RuleInfo
    {
        private static int count = 0;

        public RuleInfo()
        {
            Id = count;
            Enabled = true;
            count++;
        }

        public int Id { get; private set; }

        public int SheetIndex { get; set; }

        public int CheckSheetColumnIndex { get; set; }

        public IRowRule Rule { get; set; }

        public bool Enabled { get; set; }
    }
}
