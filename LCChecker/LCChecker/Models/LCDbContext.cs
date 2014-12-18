using LCChecker.Areas.Second.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class LCDbContext:DbContext
    {
        public LCDbContext() : base("LC") { }

        public DbSet<User> Users { get; set; }

        public DbSet<Project> Projects { get; set; }

        public DbSet<UploadFile> Files { get; set; }
       // public DbSet<CoordProjectBase> CoordProjectBase { get; set; }
        public DbSet<CoordProject> CoordProjects { get; set; }

        public DbSet<CoordNewAreaProject> CoordNewAreaProjects { get; set; }
        public DbSet<Report> Reports { get; set; }

        public DbSet<Record> Records { get; set; }
        public DbSet<SecondProject> SecondProjects { get; set; }
        public DbSet<pProject> pProjects { get; set; }
        public DbSet<SecondRecord> SecondRecords { get; set; }
        public DbSet<SecondReport> SecondReports { get; set; }
        public DbSet<FarmLand> FarmLands { get; set; }
    }
}