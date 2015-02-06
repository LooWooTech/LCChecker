

using System.Web;
namespace LCChecker.Areas.Second
{
    [UserAuthorize]
    public class SecondController : LCChecker.Controllers.ControllerBase
    {
        protected override void OnActionExecuting(System.Web.Mvc.ActionExecutingContext filterContext)
        {
            //if (!CurrentUser.Flag) {
            //    throw new HttpException(401,"你没有");
            //}
            base.OnActionExecuting(filterContext);
        }    
    }
}
