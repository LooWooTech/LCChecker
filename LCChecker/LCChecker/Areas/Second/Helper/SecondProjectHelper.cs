using LCChecker.Areas.Second.Models;
using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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

    public static class SecondProjectHelper{

        private static Regex _projectIdRe = new Regex(@"^P33[0-9]{12}", RegexOptions.Compiled);
        public static bool VerificationPID(this string value)
        {
            return _projectIdRe.IsMatch(value);
        }


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

        public static void UpDateSecondProjects(List<SecondProject> list) {
            using (var db = new LCDbContext()) {
                foreach (var item in list) {
                    var project = db.SecondProjects.Find(item.ID);
                    if (project == null)
                        continue;
                    project.Area = item.Area;
                    db.SaveChanges();
                }
                db.SaveChanges();
            }
        }

        //public static void AddPlanProject(List<pProject> list) {
        //    List<string> ID = new List<string>();
        //    using (var db = new LCDbContext()) {
        //        foreach (var item in list) {
        //            if (ID.Contains(item.ID))
        //            {
        //                continue;
        //            }
        //            else {
        //                ID.Add(item.ID);
        //            }
        //            if (!db.pProjects.Any(e => e.ID == item.ID)) {
        //                db.pProjects.Add(item);  
        //            }
        //        }
        //        db.SaveChanges();
        //    }
        //}

        public static void AddPlanAllProject(List<pProject> list) {
            using (var db = new LCDbContext()) {
                foreach (var item in list) {
                    if (!db.pProjects.Any(e => e.ID == item.ID)) {
                        db.pProjects.Add(item);
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
                if (filter.Country != null&&!string.IsNullOrEmpty(filter.Country)) {
                    query = query.Where(e => e.County.ToLower() == filter.Country.ToLower());
                }
                if (filter.Page != null) {
                    filter.Page.RecordCount = query.Count();
                    query = query.OrderBy(e => e.ID).Skip(filter.Page.PageSize * (filter.Page.PageIndex - 1)).Take(filter.Page.PageSize);
                }


                return query.ToList();
            }
        }

        public static bool CheckOnly() {
            using (var db = new LCDbContext()) {
                var All = db.pProjects.Count();
                var GNumber = db.pProjects.GroupBy(e => (e.Name + "-" + e.County + "-" + e.Key)).Select(g => g.Key).Count();
                //Dictionary<string, PlanProject> Group = db.pProjects.GroupBy(e => (e.Name + "-" + e.County + "-" + e.Key)).ToDictionary(g => g.Key, g => new PlanProject
                //{
                //    Key = g.Key,
                //    Count = g.Count()
                //});
                if (GNumber != All)
                    return false;
                else
                    return true;
            }
        }

        //20141231注释
        public static List<pProject> GetPlanProjects(SecondProjectFilter filter)
        {
            using (var db = new LCDbContext())
            {
                var query = db.pProjects.AsQueryable();
                if (filter.City.HasValue && filter.City.Value > 0)
                {
                    query = query.Where(e => e.City == filter.City.Value);
                }
                //switch (filter.Result)
                //{
                //    case NullableFilter.True:
                //    case NullableFilter.False:
                //        var value = filter.Result == NullableFilter.True;
                //        query = query.Where(e => e.Result == value);
                //        break;
                //    case NullableFilter.Null:
                //        query = query.Where(e => e.Result == null);
                //        break;
                //    case NullableFilter.All:
                //    default: break;
                //}
                if (filter.IsApplyDelete.HasValue)
                {
                    query = query.Where(e => e.IsApplyDelete == filter.IsApplyDelete.Value);
                }
                //if (filter.IsHasDoubt.HasValue)
                //{
                //    query = query.Where(e => e.IsHasDoubt == filter.IsHasDoubt.Value);
                //}

                //if (filter.IsDescrease.HasValue)
                //{
                //    query = query.Where(e => e.IsDescrease == filter.IsDescrease.Value);
                //}

                if (filter.IsHasError.HasValue)
                {
                    query = query.Where(e => e.IsHasError == filter.IsHasError.Value);
                }

                //if (filter.IsRelieve.HasValue)
                //{
                //    query = query.Where(e => e.IsRelieve == filter.IsRelieve.Value);
                //}
                //if (filter.ID != null)
                //{
                //    query = query.Where(e => e.ID.Contains(filter.ID));
                //}
                if (filter.Country != null && !string.IsNullOrEmpty(filter.Country))
                {
                    query = query.Where(e => e.County.ToLower() == filter.Country.ToLower());
                }
                if (filter.Page != null)
                {
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

        public static List<string> GetCountry(City city,bool IsPlan) {
            //List<string> Country=new List<string>();
            using (var db = new LCDbContext()) {
                if (IsPlan)
                {
                    return db.pProjects.Where(e => e.City == city).GroupBy(e => e.County).Select(g => g.Key).ToList();
                }
                else {
                    return db.SecondProjects.Where(e=>e.City==city).GroupBy(e => e.County).Select(g => g.Key).ToList();
                }

                //var list = db.SecondProjects.Where(e => e.City == city).ToList();
                //foreach (var item in list) {
                //    if (!Country.Contains(item.County)) {
                //        Country.Add(item.County);
                //    }
                    
                //}
                //return Country;
            }
        }

        public static bool Check(City city) {
            using (var db = new LCDbContext()) {
                SecondReport One = db.SecondReports.FirstOrDefault(e => e.City == city && e.IsPlan && e.Type == SecondReportType.附表1);
                if (One == null) {
                    throw new ArgumentException("请首先上传附表1之后，依次上传附表2，附表3，附表4.如有疑问请咨询管理员");
                }
                int all = 0;
                int.TryParse(One.Note, out all);
                int sum = 0;
                List<SecondReport> list = db.SecondReports.Where(e => e.City == city && e.IsPlan&&e.Type!=SecondReportType.附表1).ToList();
                foreach (var item in list) {
                    int data = 0;
                    int.TryParse(item.Note, out data);
                    sum += data;
                }
                return sum == all;
            }
        }

        public static void CheckDirectory() {
            var baseFolder = ConfigurationManager.AppSettings["BaseFolder"];
            DirectoryInfo dir = new DirectoryInfo(baseFolder);
            if (!dir.Exists) {
                dir.Create();
            }
            string file = Path.Combine(baseFolder, "Instruction.txt");
            if (!File.Exists(file)) {
                using (StreamWriter sw = File.CreateText(file))
                {
                    sw.WriteLine("本文件夹用于项目-浙江省土地整治项目核查平台使用(LCCHecker)");
                    sw.WriteLine("主要用于整理整个浙江省每个市上传的表格汇总整理时，使服务与平台保存目录识别");
                }
            }
           
            file = Path.Combine(baseFolder, "templatePath.txt");
            if (!File.Exists(file)) {
                using (StreamWriter sw = File.CreateText(file)) {
                    string txt = AppDomain.CurrentDomain.BaseDirectory;
                    sw.WriteLine(txt);
                }
            }
            //return true;
        }

    
    }

    
}