using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker
{
    public class ProjectFileter
    {
        public City? City { get; set; }

        public ProjectType? Type { get; set; }

        public NullableFilter Result { get; set; }

        public Page Page { get; set; }
    }

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

        public static List<Project> GetProjects(ProjectFileter filter)
        {
            using (var db = new LCDbContext())
            {
                var query = db.Projects.AsQueryable();
                if (filter.City.HasValue && filter.City.Value > 0)
                {
                    query = query.Where(e => e.City == filter.City.Value);
                }

                switch (filter.Result)
                {
                    case NullableFilter.True:
                    case NullableFilter.False:
                        var value = filter.Result == NullableFilter.True;
                        query = query.Where(e => e.Result == value);
                        break;
                    case NullableFilter.Null:
                        query = query.Where(e => e.Result == null);
                        break;
                    case NullableFilter.All:
                    default:
                        break;
                }

                if (filter.Type.HasValue && filter.Type.Value > 0)
                {
                    query = query.Where(e => e.Type == filter.Type.Value);
                }

                if (filter.Page != null)
                {
                    filter.Page.RecordCount = query.Count();
                    query = query.OrderByDescending(e => e.UpdateTime).Skip(filter.Page.PageSize * (filter.Page.PageIndex - 1)).Take(filter.Page.PageSize);
                }

                return query.ToList();
            }
        }
    }
}