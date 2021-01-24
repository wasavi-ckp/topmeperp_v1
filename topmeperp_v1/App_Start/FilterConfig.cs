using System.Web;
using System.Web.Mvc;
using topmeperp.Filter;

namespace topmeperp
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new AuthFilter());
            filters.Add(new HandleErrorAttribute());
        }
    }
}
