using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class Page
    {
        public Page()
        {
            PageSize = 40;
            PageIndex = 1;
        }

        public Page(int page)
            : this()
        {
            PageIndex = page < 1 ? 1 : page;
        }

        public int RecordCount { get; set; }

        public int PageSize { get; set; }

        public int PageIndex { get; set; }

        public int PageCount
        {
            get
            {
                var count = RecordCount / PageSize;
                var last = RecordCount % PageSize;
                return last > 0 ? count++ : count;
            }
        }
    }
}