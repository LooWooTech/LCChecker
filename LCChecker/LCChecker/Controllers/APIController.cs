using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LCChecker.Controllers
{
    public class ProjectJsonModel
    {
        public string id { get; set; }

        public bool result { get; set; }

        public string error { get; set; }
    }

    public class APIController : ControllerBase
    {
        private System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();

        private ActionResult ErrorResult(string message)
        {
            return Json(new { result = 0, message });
        }

        public ActionResult UpdateProjects(City? city = null)
        {
            //{id:"xxxxx",result:true|false,error:""}
            var json = Request["data"];
            if (string.IsNullOrEmpty(json))
            {
                return ErrorResult("data参数为空");
            }

            try
            {
                var data = serializer.Deserialize<List<ProjectJsonModel>>(json);

                foreach (var item in data)
                {
                    var entity = db.Projects.FirstOrDefault(e => e.ID == item.id);
                    if (entity != null)
                    {
                        entity.Result = item.result;
                        entity.Note = item.error;
                    }
                    db.SaveChanges();
                }
                return Json(new { result = 1 });
            }
            catch (Exception ex)
            {
                return ErrorResult(ex.Message);
            }
        }
    }
}
