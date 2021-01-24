using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace topmeperp.Controllers
{
    public class TestController : Controller
    {
        // GET: Test
        public ActionResult Index()
        {
            return View();
        }
        //handsontable : Excel 樣式範例
        public ActionResult handsontable()
        {
            return View();
        }
    }
    
}