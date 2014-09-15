using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker
{
    public class ProjectHelper
    {
        public static void AddProjects(List<Project> list)
        {
            using (var db = new LCDbContext())
            {
                foreach (var item in list)
                {
                    if (!db.Projects.Any(e => e.ID == item.ID))
                    {
                        db.Projects.Add(item);
                    }
                    db.SaveChanges();
                }
            }
        }

        public static List<Project> GetProjects(City city = City.浙江省, ResultFilter result = ResultFilter.All, Page page = null)
        {
            using (var db = new LCDbContext())
            {
                bool? resultValue = null;
                switch (result)
                {
                    case ResultFilter.Pass:
                        resultValue = true;
                        break;
                    case ResultFilter.Error:
                        resultValue = false;
                        break;
                    case ResultFilter.Uncheck:
                        resultValue = null;
                        break;
                }

                var query = db.Projects.AsQueryable();
                if (city != City.浙江省)
                {
                    query = query.Where(e => e.City == city);
                }

                if (result != ResultFilter.All)
                {
                    query = query.Where(e => e.Result == resultValue);
                }

                if (page != null)
                {
                    page.RecordCount = query.Count();
                    query = query.OrderBy(e => e.ID).Skip(page.PageSize * (page.PageIndex - 1)).Take(page.PageSize);
                }

                return query.ToList();
            }
        }
    }
}