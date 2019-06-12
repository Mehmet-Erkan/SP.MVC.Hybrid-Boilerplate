using System.Web;
using System.Web.Mvc;

namespace SP.MVC.ReactBoilerplateWeb
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
