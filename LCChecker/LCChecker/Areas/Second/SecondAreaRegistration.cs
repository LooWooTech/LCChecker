using System.Web.Mvc;

namespace LCChecker.Areas.Second
{
    public class SecondAreaRegistration : AreaRegistration
    {
        public override string AreaName
        {
            get
            {
                return "Second";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context)
        {
            context.MapRoute(
                "Second_default",
                "Second/{controller}/{action}/{id}",
                new { controller="Home", action = "Index", id = UrlParameter.Optional },
                new string[] { "LCChecker.Areas.Second.Controllers"}
            );
        }
    }
}
