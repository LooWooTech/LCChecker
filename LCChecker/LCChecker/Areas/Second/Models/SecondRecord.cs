using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace LCChecker.Areas.Second.Models
{
    [Table("serecords")]
    public class SecondRecord
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }
        [MaxLength(55)]
        public string ProjectID { get; set; }
        [Column(TypeName="int")]
        public SecondReportType Type { get; set; }
        [Column(TypeName = "int")]
        public City City { get; set; }

        public bool IsError { get; set; }
        [MaxLength(1023)]
        public string Note { get; set; }



        public static void Clear(List<SecondRecord> List) {
            using (var db = new LCDbContext()) {
                foreach (var item in List) {
                    db.SecondRecords.Attach(item);
                    db.SecondRecords.Remove(item);
                }
                db.SaveChanges();
            }
        }

        public static void AddRecords(List<SecondRecord> List) {
            using (var db = new LCDbContext()) {
                foreach (var item in List) {
                    db.SecondRecords.Add(item);
                }
                try
                {
                    db.SaveChanges();
                }
                catch { 
                
                }
            }
        }
    }
}