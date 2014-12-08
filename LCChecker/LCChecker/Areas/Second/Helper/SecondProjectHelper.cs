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
        public string ID { get; set; }
        public string Country { get; set; }
        public Page Page { get; set; }
        public bool? Vieible { get; set; }
        public bool? IsHasDoubt { get; set; }
        public bool? IsHasError { get; set; }
        public bool? IsApplyDelete { get; set; }
        public bool? IsDescrease { get; set; }
        public bool? IsRelieve { get; set; }
    }




    public class SeProject {
        public bool IsHasDoubt { get; set; }
        public bool IsApplyDelete { get; set; }
        public bool IsHasError { get; set; }
        public bool IsPacket { get; set; }
        public bool IsDescrease { get; set; }
        public bool IsRelieve { get; set; }
    }


    public class SeLand {
        public Degree Degree { get; set; }
        public double Area { get; set; }
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

        public static void AddCoordNewArea(List<CoordNewAreaProject> list) {
            using (var db = new LCDbContext()) {
                foreach (var item in list) {
                    if (!db.CoordNewAreaProjects.Any(e => e.ID == item.ID)) {
                        db.CoordNewAreaProjects.Add(item);
                    }
                }
                db.SaveChanges();
            }
        }


        public static void AddCoordProject(List<CoordProject> list) {
            using (var db = new LCDbContext()) {
                foreach (var item in list) {
                    if (!db.CoordProjects.Any(e => e.ID == item.ID)) {
                        db.CoordProjects.Add(item);
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
                if (filter.ID!=null) {
                    query = query.Where(e => e.ID.Contains(filter.ID));
                }
                if (filter.Country != null) {
                    query = query.Where(e => e.County.ToLower() == filter.Country.ToLower());
                }
                if (filter.Page != null) {
                    filter.Page.RecordCount = query.Count();
                    query = query.OrderBy(e => e.ID).Skip(filter.Page.PageSize * (filter.Page.PageIndex - 1)).Take(filter.Page.PageSize);
                }


                return query.ToList();
            }
        }
        public static List<CoordNewAreaProject> GetNewAreaCoord(SecondProjectFilter Filter) {
            using (var db = new LCDbContext()) {
                var query = db.CoordNewAreaProjects.AsQueryable();
                if (Filter.City.HasValue && Filter.City.Value > 0) {
                    query = query.Where(e => e.City == Filter.City.Value);
                }
                switch (Filter.Result) { 
                    case NullableFilter.True:
                    case NullableFilter.False:
                        var value = Filter.Result == NullableFilter.True;
                        query = query.Where(e => e.Result == value);
                        break;
                    case NullableFilter.Null:
                        query = query.Where(e => e.Result == null);
                        break;
                    case NullableFilter.All:
                    default: break;
                }

                if (Filter.Page != null) {
                    Filter.Page.RecordCount = query.Count();
                    query = query.OrderBy(e => e.ID).Skip(Filter.Page.PageSize * (Filter.Page.PageIndex - 1)).Take(Filter.Page.PageSize);
                }

                return query.ToList();
            }
            
        }

        public static List<string> GetCountry(City city) {
            List<string> Country=new List<string>();
            using (var db = new LCDbContext()) {
                var list = db.SecondProjects.Where(e => e.City == city).ToList();
                foreach (var item in list) {
                    if (!Country.Contains(item.County)) {
                        Country.Add(item.County);
                    }
                    
                }
                return Country;
            }
        }
    
    }

    
}