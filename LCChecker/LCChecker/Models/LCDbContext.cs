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
        public DbSet<User> USER { get; set; }
        public DbSet<Detect> DETECT { get; set; }
        public DbSet<SubRecord> SUBRECORD { get; set; }
    }
}