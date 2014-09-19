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

        public NullableFilter Result { get; set; }

        public Page Page { get; set; }

        public bool? Visible { get; set; }

        public bool? IsShouldModify { get; set; }

        public bool? IsApplyDelete { get; set; }

        public bool? IsHasError { get; set; }

        public bool? IsDecrease { get; set; }
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
                }
                db.SaveChanges();
            }
        }

        public static void AddCoordProjects(List<CoordProject> list)
        {
            using (var db = new LCDbContext())
            {
                foreach (var item in list)
                {
                    var entity = db.CoordProjects.FirstOrDefault(e => e.ID == item.ID);
                    if (entity == null)
                    {
                        db.CoordProjects.Add(item);
                    }
                    else if(item.Visible == true)
                    {
                        entity.UpdateTime = DateTime.Now;
                        entity.Note = item.Note;
                        entity.Visible = item.Visible;
                    }
                }
                db.SaveChanges();
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


                if (filter.IsApplyDelete.HasValue)
                {
                    query = query.Where(e => e.IsApplyDelete == filter.IsApplyDelete.Value);
                }

                if (filter.IsDecrease.HasValue)
                {
                    query = query.Where(e => e.IsDecrease == filter.IsDecrease.Value);
                }

                if (filter.IsHasError.HasValue)
                {
                    query = query.Where(e => e.IsHasError == filter.IsHasError.Value);
                }

                if (filter.IsShouldModify.HasValue)
                {
                    query = query.Where(e => e.IsShouldModify == filter.IsShouldModify.Value);
                }

                if (filter.Page != null)
                {
                    filter.Page.RecordCount = query.Count();
                    query = query.OrderBy(e => e.ID).Skip(filter.Page.PageSize * (filter.Page.PageIndex - 1)).Take(filter.Page.PageSize);
                }

                return query.ToList();
            }
        }

        public static List<CoordProject> GetCoordProjects(ProjectFileter filter)
        {
            using (var db = new LCDbContext())
            {
                var query = db.CoordProjects.AsQueryable();
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

                if (filter.Visible.HasValue)
                {
                    query = query.Where(e => e.Visible == filter.Visible.Value);
                }

                //if (filter.Type.HasValue && filter.Type.Value > 0)
                //{
                //    query = query.Where(e => e.Type == filter.Type.Value);
                //}

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