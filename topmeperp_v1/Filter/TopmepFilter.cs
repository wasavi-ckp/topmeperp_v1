using log4net;
using System.Collections.Generic;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using topmeperp.Models;

namespace topmeperp.Filter
{
    public class AuthFilter : ActionFilterAttribute
    {
        static ILog log = log4net.LogManager.GetLogger(typeof(AuthFilter));
        public override void OnActionExecuting(ActionExecutingContext context)
        {
            log.Info("request URL=" + context.HttpContext.Request.RawUrl);
            if (context.HttpContext.Session["user"] != null || exceptionUrl(context.HttpContext.Request.RawUrl))
            {
                //驗證成功
                log.Info("session exist!!");
            }
            else
            {
                log.Info("session not exist!!");
                if (context.HttpContext.Request.RawUrl != "/"
                    && context.HttpContext.Request.RawUrl != "/topmeperp/")
                {
                    log.Info("forward to login page");
                    context.Result = new RedirectToRouteResult(
                    new RouteValueDictionary
                    {
                         { "controller", "Home" },
                         { "action", "Login" }
                    });
                }
            }

        }
        public bool exceptionUrl(string url)
        {
            string[] exceptinUrl = new string[] {""};
            for (int i=0;i< exceptinUrl.Length; i++)
            {
                if (url.Equals(exceptinUrl[i]))
                {
                    return true;
                }
            }
            return false;
        }
        public static string getBreadcrumb(string url)
        {
            string breadcrumbHtml = "";
            List<SYS_FUNCTION> lst = (List<SYS_FUNCTION>)HttpContext.Current.Session["functions"];
            //使用list 物件查詢功能
            SYS_FUNCTION curFunction = lst.Find(x => x.FUNCTION_URI.Contains(url));
            log.Info("cur url=" + curFunction);
            if (null != curFunction && curFunction.ISMENU == "Y")
            {
                //未來有樹狀較完整後再調整
                //string[] breadcrumb = url.Split('/');
                // for (int i = 1; i < breadcrumb.Length; i++)
                // {
                //breadcrumbHtml = breadcrumbHtml + "<li class='breadcrumb-item'><a href='"+ curFunction.FUNCTION_URI + "'>"+ curFunction.FUNCTION_NAME+ "</a></li>";
                // }
                breadcrumbHtml = breadcrumbHtml + "<li class='breadcrumb-item'><a href='#'>" + curFunction.MODULE_NAME + "</a></li>";
                breadcrumbHtml = breadcrumbHtml + "<li class='breadcrumb-item'><a href='" + curFunction.FUNCTION_URI + "'>" + curFunction.FUNCTION_NAME + "</a></li>";
                HttpContext.Current.Session["sitepath"]= breadcrumbHtml;
            }else
            {
                breadcrumbHtml = (string)HttpContext.Current.Session["sitepath"];
            }
            return breadcrumbHtml;
        }
    }

}