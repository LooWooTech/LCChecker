using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second
{
    public class SecondProjectFilter
    {
        public City? City { get; set; }
        public NullableFilter Result { get; set; }
        public Page Page { get; set; }
        public bool? Vieible { get; set; }
        public bool? IsHasDoubt { get; set; }
        public bool? IsHasError { get; set; }
        public bool? IsApplyDelete { get; set; }
        public bool? IsDescrease { get; set; }
        public bool? IsRelieve { get; set; }
    }

    public class SecondProjectHelper{
        public static void AddSecondProjects(List<SecondProject> list)
        {
            using (var db = new LCDbContext())
            {
                foreach (var item in list)
                {
                    if (!db.SecondProjects.Any(e => e.ID == item.ID))
                    {
                        db.SecondProjects.Add(item);
                    }
                }
                db.SaveChanges();
            }
        }

        public static List<SecondProject> GetProjects(SecondProjectFilter filter) {
            using (var db = new LCDbContext()) {
                var query = db.SecondProjects.AsQueryable();
                if (filter.City.HasValue && filter.City.Value > 0) {
                    query = query.Where(e => e.City == filter.City.Value);
                }
                switch (filter.Result) { 
                    case NullableFilter.True:
                    case NullableFilter.False:
                        var value = filter.Result == NullableFilter.True;
                        query = query.Where(e => e.Result == value);
                        break;
                    case NullableFilter.Null:
                        query = query.Where(e => e.Result == null);
                        break;
                    case NullableFilter.All:
                    default: break;
                }
                if (filter.IsApplyDelete.HasValue) {
                    query = query.Where(e => e.IsApplyDelete == filter.IsApplyDelete.Value);
                }
                if (filter.IsHasDoubt.HasValue) {
                    query = query.Where(e => e.IsHasDoubt == filter.IsHasDoubt.Value);
                }

                if (filter.IsDescrease.HasValue)
                {
                    query = query.Where(e => e.IsDescrease == filter.IsDescrease.Value);
                }

                if (filter.IsHasError.HasValue) {
                    query = query.Where(e => e.IsHasError == filter.IsHasError.Value);
                }

                if (filter.IsRelieve.HasValue) {
                    query = query.Where(e => e.IsRelieve == filter.IsRelieve.Value);
                }
                if (filter.Page != null) {
                    filter.Page.RecordCount = query.Count();
                    query = query.OrderBy(e => e.ID).Skip(filter.Page.PageSize * (filter.Page.PageIndex - 1)).Take(filter.Page.PageSize);
                }


                return query.ToList();
            }
        }
    
    }

    
}