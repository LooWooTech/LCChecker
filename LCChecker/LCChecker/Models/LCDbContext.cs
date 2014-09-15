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
    }
}